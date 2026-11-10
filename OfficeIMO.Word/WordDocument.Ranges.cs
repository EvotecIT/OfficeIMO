using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Internal;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        /// <summary>Text replacements made through ranges, in order; other ranges replay them to stay on their text.</summary>
        internal List<WordRangeEdit> RangeEdits { get; } = new();

        /// <summary>
        /// Replaces regular-expression matches in the body, body tables, all section headers and footers, and ordinary
        /// footnotes and endnotes.
        /// </summary>
        /// <param name="pattern">Regular expression applied to each paragraph's text.</param>
        /// <param name="replacement">
        /// Replacement pattern with the substitutions of <see cref="Regex.Replace(string, string)"/>, such as <c>$1</c> or <c>${name}</c>.
        /// </param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>
        /// Each match lies within one paragraph and may span runs; the new text takes the formatting of the first matched character.
        /// When any match cannot be edited safely, such as a field result or a tracked change, the method throws
        /// <see cref="InvalidOperationException"/> before replacing anything.
        /// </remarks>
        public int ReplaceAll(Regex pattern, string replacement) {
            ArgumentNullException.ThrowIfNull(replacement);
            return ReplaceAll(pattern, match => match.Result(replacement));
        }

        /// <summary>
        /// Replaces regular-expression matches with text computed for each match, in the same parts as <see cref="ReplaceAll(Regex, string)"/>.
        /// </summary>
        /// <param name="pattern">Regular expression applied to each paragraph's text.</param>
        /// <param name="evaluator">Returns the replacement text for a match.</param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>Matching, formatting and safety checks follow <see cref="ReplaceAll(Regex, string)"/>.</remarks>
        public int ReplaceAll(Regex pattern, MatchEvaluator evaluator) {
            ArgumentNullException.ThrowIfNull(evaluator);
            return ReplaceMatches(pattern, (range, match) => range.Replace(evaluator(match) ?? string.Empty));
        }

        /// <summary>
        /// Replaces regular-expression matches with content added through a paragraph, such as formatted text, a hyperlink
        /// or an image, in the same parts as <see cref="ReplaceAll(Regex, string)"/>.
        /// </summary>
        /// <param name="pattern">Regular expression applied to each paragraph's text.</param>
        /// <param name="content">
        /// Adds the replacement for a match, for example <c>(m, p) =&gt; p.AddText(m.Groups[1].Value).SetBold()</c>.
        /// </param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>Matching and safety checks follow <see cref="ReplaceAll(Regex, string)"/>.</remarks>
        public int ReplaceAll(Regex pattern, Action<Match, WordParagraph> content) {
            ArgumentNullException.ThrowIfNull(content);
            return ReplaceMatches(pattern, (range, match) => range.Replace(paragraph => content(match, paragraph)));
        }

        private int ReplaceMatches(Regex pattern, Action<WordRange, Match> replace) {
            ArgumentNullException.ThrowIfNull(pattern);
            List<(WordRange Range, Match Match)> matches = FindRegexRanges(pattern, null);
            foreach (var (range, _) in matches) range.ValidateReplace();
            // Replacing from the end keeps the offsets of earlier matches in the same paragraph unchanged.
            for (int index = matches.Count - 1; index >= 0; index--) {
                replace(matches[index].Range, matches[index].Match);
            }
            return matches.Count;
        }

        internal List<WordRange> FindTextRanges(string text, OfficeTextSearchOptions? options, WordSection? section) {
            StringComparison comparison = OfficeTextSearch.GetComparison(options);
            bool wholeWords = options?.WholeWords == true;
            var ranges = new List<WordRange>();
            foreach (List<WordTextMap> story in GetSearchStories(section)) {
                if (wholeWords) {
                    foreach (WordTextMap map in story) {
                        int from = 0;
                        while (from <= map.Text.Length - text.Length) {
                            int found = map.Text.IndexOf(text, from, comparison);
                            if (found < 0) break;
                            if (!OfficeTextSearch.HasWordBoundaries(map.Text, found, text.Length)) {
                                from = found + 1;
                                continue;
                            }
                            ranges.Add(new WordRange(this, map.Paragraph, found, map.Paragraph, found + text.Length));
                            from = found + text.Length;
                        }
                    }
                    continue;
                }

                // Without whole-word matching a match may continue into the next paragraph of the same story.
                string combined = string.Concat(story.Select(map => map.Text));
                int searchFrom = 0;
                while (searchFrom <= combined.Length - text.Length) {
                    int found = combined.IndexOf(text, searchFrom, comparison);
                    if (found < 0) break;
                    (WordTextMap startMap, int start) = Locate(story, found, atEnd: false);
                    (WordTextMap endMap, int end) = Locate(story, found + text.Length, atEnd: true);
                    ranges.Add(new WordRange(this, startMap.Paragraph, start, endMap.Paragraph, end));
                    searchFrom = found + text.Length;
                }
            }
            return ranges;
        }

        private static (WordTextMap Map, int Offset) Locate(List<WordTextMap> story, int offset, bool atEnd) {
            int paragraphStart = 0;
            foreach (WordTextMap map in story) {
                int paragraphEnd = paragraphStart + map.Text.Length;
                if (atEnd ? offset <= paragraphEnd && offset > paragraphStart : offset < paragraphEnd) {
                    return (map, offset - paragraphStart);
                }
                paragraphStart = paragraphEnd;
            }
            WordTextMap last = story[story.Count - 1];
            return (last, last.Text.Length);
        }

        internal List<(WordRange Range, Match Match)> FindRegexRanges(Regex pattern, WordSection? section) {
            var ranges = new List<(WordRange, Match)>();
            foreach (List<WordTextMap> story in GetSearchStories(section)) {
                foreach (WordTextMap map in story) {
                    foreach (Match match in pattern.Matches(map.Text)) {
                        ranges.Add((new WordRange(this, map.Paragraph, match.Index, map.Paragraph, match.Index + match.Length), match));
                    }
                }
            }
            return ranges;
        }

        /// <summary>The paragraphs of each search scope, as visited by <c>Find</c> and <c>ReplaceAll</c>.</summary>
        private List<List<WordTextMap>> GetSearchStories(WordSection? section) {
            var stories = new List<List<WordTextMap>>();
            ForEachSearchScope((paragraphs, _) => {
                var story = new List<WordTextMap>();
                Paragraph? previous = null;
                foreach (WordParagraph paragraph in paragraphs) {
                    if (paragraph == null || ReferenceEquals(paragraph._paragraph, previous)) continue;
                    previous = paragraph._paragraph;
                    story.Add(WordTextMap.Build(previous));
                }
                if (story.Count > 0) stories.Add(story);
            }, null, section);
            return stories;
        }

        /// <summary>Returns an unused bookmark identifier across the body, headers, footers, notes and comments.</summary>
        internal int GetNextBookmarkId() {
            int max = 0;
            foreach (OpenXmlElement root in EnumerateStoryRoots()) {
                foreach (BookmarkStart start in root.Descendants<BookmarkStart>()) {
                    if (int.TryParse(start.Id?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) && id > max) max = id;
                }
            }
            return max + 1;
        }

        private IEnumerable<OpenXmlElement> EnumerateStoryRoots() {
            MainDocumentPart? main = _wordprocessingDocument?.MainDocumentPart;
            if (main == null) yield break;
            if (main.Document?.Body != null) yield return main.Document.Body;
            foreach (HeaderPart header in main.HeaderParts) if (header.Header != null) yield return header.Header;
            foreach (FooterPart footer in main.FooterParts) if (footer.Footer != null) yield return footer.Footer;
            if (main.FootnotesPart?.Footnotes != null) yield return main.FootnotesPart.Footnotes;
            if (main.EndnotesPart?.Endnotes != null) yield return main.EndnotesPart.Endnotes;
            if (main.WordprocessingCommentsPart?.Comments != null) yield return main.WordprocessingCommentsPart.Comments;
        }
    }
}
