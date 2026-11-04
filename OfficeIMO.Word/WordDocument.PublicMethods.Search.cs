using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Http;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private int CombineRuns(WordHeaderFooter wordHeaderFooter) {
            int count = 0;
            if (wordHeaderFooter != null) {
                var defaultHeader = this.Sections[0].Headers?.Default;
                if (defaultHeader != null) {
                    foreach (var p in defaultHeader.Paragraphs) count += CombineIdenticalRuns(p._paragraph);
                    foreach (var table in defaultHeader.Tables) {
                        table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                    }
                }
            }

            return count;
        }


        /// <summary>
        /// This method will combine identical runs in a paragraph.
        /// This is useful when you have a paragraph with multiple runs of the same style, that Microsoft Word creates.
        /// This feature is *EXPERIMENTAL* and may not work in all cases.
        /// It may impact on how your document looks like, please do extensive testing before using this feature.
        /// </summary>
        /// <returns></returns>
        public int CleanupDocument(WordDocumentCleanupOptions options = WordDocumentCleanupOptions.All) {
            int count = 0;

            if (_wordprocessingDocument?.MainDocumentPart?.Document?.Body != null) {
                foreach (var paragraph in _wordprocessingDocument.MainDocumentPart.Document.Body.Descendants<Paragraph>().ToList()) {
                    count += CleanupParagraph(paragraph, options);
                }
            }

            foreach (var header in _wordprocessingDocument?.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>()) {
                foreach (var paragraph in (header.Header?.Descendants<Paragraph>() ?? Enumerable.Empty<Paragraph>()).ToList()) {
                    count += CleanupParagraph(paragraph, options);
                }
            }

            foreach (var footer in _wordprocessingDocument?.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>()) {
                foreach (var paragraph in (footer.Footer?.Descendants<Paragraph>() ?? Enumerable.Empty<Paragraph>()).ToList()) {
                    count += CleanupParagraph(paragraph, options);
                }
            }

            return count;
        }

        /// <summary>
        /// Finds literal text in the body, body tables, and the first section's headers and footers.
        /// </summary>
        /// <param name="text">Text to find.</param>
        /// <param name="options">Comparison settings. <c>null</c> matches ordinally, ignoring case, inside or across words.</param>
        /// <returns>The paragraphs containing matches, grouped by document part, and the total match count.</returns>
        /// <remarks>
        /// A match may span runs within a paragraph. Without <see cref="OfficeTextSearchOptions.WholeWords"/>, a match may
        /// also continue into the next paragraph of the same part; its result entry is the paragraph where it starts.
        /// With <see cref="OfficeTextSearchOptions.WholeWords"/>, each match lies within one paragraph.
        /// </remarks>
        public WordSearchResult Find(string text, OfficeTextSearchOptions? options = null) {
            OfficeTextSearch.ValidateSearchText(text, nameof(text));
            var result = new WordSearchResult();
            StringComparison comparison = OfficeTextSearch.GetComparison(options);
            bool wholeWords = options?.WholeWords == true;
            int count = 0;
            ForEachSearchScope((paragraphs, destination) =>
                destination!.AddRange(ReplaceText(paragraphs, text, string.Empty, ref count, false, comparison, wholeWords)), result);
            result.MatchCount = count;
            return result;
        }

        /// <summary>
        /// Finds text matching a regular expression in the body, body tables, and the first section's headers and footers.
        /// </summary>
        /// <param name="pattern">Regular expression applied to each paragraph's text. Its own options control case and boundaries.</param>
        /// <returns>The paragraphs containing matches, grouped by document part, and the total match count.</returns>
        /// <remarks>Each match lies within one paragraph.</remarks>
        public WordSearchResult Find(Regex pattern) {
            if (pattern == null) {
                throw new ArgumentNullException(nameof(pattern));
            }

            var result = new WordSearchResult();
            ForEachSearchScope((paragraphs, destination) => result.FindRegex(paragraphs, pattern, destination!), result);
            return result;
        }

        /// <summary>
        /// Replaces literal text in the body, body tables, and the first section's headers and footers.
        /// </summary>
        /// <param name="text">Text to find.</param>
        /// <param name="replacement">Replacement text; <c>null</c> removes each match.</param>
        /// <param name="options">Comparison settings. <c>null</c> matches ordinally, ignoring case, inside or across words.</param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>
        /// A match may span runs within a paragraph.
        /// Without <see cref="OfficeTextSearchOptions.WholeWords"/>, a match continuing into the next paragraph joins
        /// those paragraphs. With <see cref="OfficeTextSearchOptions.WholeWords"/>, each match lies within one paragraph.
        /// </remarks>
#pragma warning disable RS0026 // Document and paragraph-scoped ReplaceAll overloads both keep the options parameter optional.
        public int ReplaceAll(string text, string replacement, OfficeTextSearchOptions? options = null) {
#pragma warning restore RS0026
            OfficeTextSearch.ValidateSearchText(text, nameof(text));
            StringComparison comparison = OfficeTextSearch.GetComparison(options);
            bool wholeWords = options?.WholeWords == true;
            if (!wholeWords && TryFindAndReplaceInSimpleTextNodes(text, replacement, comparison, out int fastCount)) {
                return fastCount;
            }
            int count = 0;
            ForEachSearchScope((paragraphs, _) => ReplaceText(paragraphs, text, replacement, ref count, true, comparison, wholeWords), null);
            return count;
        }

        private bool TryFindAndReplaceInSimpleTextNodes(
            string textToFind,
            string textToReplace,
            StringComparison stringComparison,
            out int count) {
            count = 0;
            string replacement = textToReplace ?? string.Empty;
            if (string.IsNullOrEmpty(textToFind) ||
                replacement.IndexOf('\r') >= 0 ||
                replacement.IndexOf('\n') >= 0 ||
                replacement.IndexOf('\t') >= 0) {
                return false;
            }

            MainDocumentPart? mainPart = _wordprocessingDocument?.MainDocumentPart;
            Body? body = mainPart?.Document?.Body;
            if (body == null || mainPart!.HeaderParts.Any() || mainPart.FooterParts.Any()) {
                return false;
            }

            var nodes = new List<Text>();
            OpenXmlElement? bodyChild = body.FirstChild;
            while (bodyChild != null) {
                if (bodyChild is Paragraph paragraph) {
                    if (!TryCollectSimpleTextNodes(paragraph, nodes)) return false;
                } else if (bodyChild is not SectionProperties) {
                    return false;
                }
                bodyChild = bodyChild.NextSibling();
            }
            string[] values = new string[nodes.Count];
            for (int index = 0; index < nodes.Count; index++) values[index] = WordParagraph.ReadVisibleText(nodes[index]);
            if (HasCrossParagraphMatch(values, textToFind, stringComparison)) return false;

            for (int index = 0; index < nodes.Count; index++) {
                string value = values[index];
                if (value.IndexOf(textToFind, stringComparison) < 0) continue;
                int localCount = 0;
                string updated = value.FindAndReplace(
                    textToFind,
                    replacement,
                    stringComparison,
                    ref localCount);
                nodes[index].Text = updated;
                if (updated.Length > 0 &&
                    (char.IsWhiteSpace(updated[0]) || char.IsWhiteSpace(updated[updated.Length - 1]))) {
                    nodes[index].Space = SpaceProcessingModeValues.Preserve;
                }
                count += localCount;
            }
            return true;
        }

        private static bool TryCollectSimpleTextNodes(OpenXmlElement parent, List<Text> nodes) {
            OpenXmlElement? child = parent.FirstChild;
            while (child != null) {
                if (child is Text text) {
                    nodes.Add(text);
                } else if (child is OpenXmlLeafTextElement ||
                           child is Break ||
                           child is TabChar ||
                           child is CarriageReturn) {
                    return false;
                } else if (child.HasChildren && !TryCollectSimpleTextNodes(child, nodes)) {
                    return false;
                }
                child = child.NextSibling();
            }
            return true;
        }

        /// <summary>
        /// Replaces literal text within the supplied paragraphs, such as the paragraphs of one table cell.
        /// </summary>
        /// <param name="paragraphs">Paragraphs to search, in document order.</param>
        /// <param name="text">Text to find.</param>
        /// <param name="replacement">Replacement text; <c>null</c> removes each match.</param>
        /// <param name="options">Comparison settings. <c>null</c> matches ordinally, ignoring case, inside or across words.</param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>Matching follows <see cref="ReplaceAll(string, string, OfficeTextSearchOptions?)"/>, limited to <paramref name="paragraphs"/>.</remarks>
#pragma warning disable RS0026 // Paragraph-scoped overload of the document ReplaceAll; both keep the options parameter optional.
        public static int ReplaceAll(List<WordParagraph> paragraphs, string text, string replacement, OfficeTextSearchOptions? options = null) {
#pragma warning restore RS0026
            if (paragraphs == null) {
                throw new ArgumentNullException(nameof(paragraphs));
            }
            int count = 0;
            ReplaceText(paragraphs, text, replacement, ref count, true, OfficeTextSearch.GetComparison(options), options?.WholeWords == true);
            return count;
        }

        /// <summary>
        /// Replace text inside each paragraph
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="oldText"></param>
        /// <param name="newText"></param>
        /// <param name="count"></param>
        /// <param name="replace"></param>
        /// <param name="stringComparison"></param>
        /// <param name="wholeWords">Requires word boundaries and keeps each match within one paragraph.</param>
        /// <returns></returns>
        private static List<WordParagraph> ReplaceText(List<WordParagraph> paragraphs, string oldText, string? newText, ref int count, bool replace, StringComparison stringComparison, bool wholeWords) {
            OfficeTextSearch.ValidateSearchText(oldText, "text");
            newText ??= string.Empty;
            List<WordParagraph> foundParagraphs = new List<WordParagraph>();
            string[] paragraphTexts = new string[paragraphs.Count];
            for (int index = 0; index < paragraphs.Count; index++) {
                paragraphTexts[index] = paragraphs[index]?.Text ?? string.Empty;
            }
            if (wholeWords || !HasCrossParagraphMatch(paragraphTexts, oldText, stringComparison)) {
                for (int index = 0; index < paragraphs.Count; index++) {
                    WordParagraph? paragraph = paragraphs[index];
                    if (paragraph == null) continue;
                    int localCount = 0;
                    string updated = OfficeTextSearch.Replace(
                        paragraphTexts[index],
                        oldText,
                        newText,
                        stringComparison,
                        wholeWords,
                        ref localCount);
                    if (localCount == 0) continue;

                    count += localCount;
                    if (replace) paragraph.Text = updated;
                    foundParagraphs.Add(paragraph);
                }
                return foundParagraphs;
            }

            var removeParas = new List<int>();
            var foundList = SearchText(paragraphs, oldText, new WordPositionInParagraph() { Paragraph = 0 }, stringComparison);

            if (foundList?.Count > 0) {
                count += foundList.Count;
                foreach (var ts in foundList) {
                    if (!IsSegmentValid(paragraphs, ts))
                        continue;
                    if (ts.BeginIndex == ts.EndIndex) {
                        var p = paragraphs[ts.BeginIndex];
                        if (p is not null) {
                            if (replace) {
                                int replaceCount = 0;
                                p.Text = p.Text.FindAndReplace(oldText, newText, stringComparison, ref replaceCount);
                            }
                            if (!foundParagraphs.Any(fp => ReferenceEquals(fp._paragraph, p._paragraph))) {
                                foundParagraphs.Add(p);
                            }
                        }
                    } else {
                        if (!replace && !foundParagraphs.Any(fp => ReferenceEquals(fp._paragraph, paragraphs[ts.BeginIndex]._paragraph))) {
                            foundParagraphs.Add(paragraphs[ts.BeginIndex]);
                        }
                        if (replace) {
                            var beginPara = paragraphs[ts.BeginIndex];
                            var endPara = paragraphs[ts.EndIndex];
                            if (beginPara is not null && endPara is not null) {
                                beginPara.Text = beginPara.Text.Replace(beginPara.Text.Substring(ts.BeginChar), newText);
                                endPara.Text = endPara.Text.Replace(endPara.Text.Substring(0, ts.EndChar + 1), "");
                                if (!foundParagraphs.Any(fp => ReferenceEquals(fp._paragraph, beginPara._paragraph))) {
                                    foundParagraphs.Add(beginPara);
                                }
                            }
                            for (int i = ts.EndIndex - 1; i > ts.BeginIndex; i--) {
                                removeParas.Add(i);
                            }
                        }

                    }
                }
            }

            if (replace) {
                if (removeParas.Count > 0) {
                    removeParas = removeParas.Distinct().OrderByDescending(i => i).ToList();// Need remove by descending
                    foreach (var index in removeParas) {
                        paragraphs[index].Remove();//Remove blank paragraph
                    }
                }
            }
            return foundParagraphs;
        }

        private static bool HasCrossParagraphMatch(
            IReadOnlyList<string> paragraphTexts,
            string searched,
            StringComparison stringComparison) {
            if (searched.Length < 2 || paragraphTexts.Count < 2) return false;
            if (stringComparison == StringComparison.Ordinal) {
                return HasCrossParagraphOrdinalMatch(paragraphTexts, searched);
            }

            int tailLength = searched.Length - 1;
            string tail = paragraphTexts[0].Length <= tailLength
                ? paragraphTexts[0]
                : paragraphTexts[0].Substring(paragraphTexts[0].Length - tailLength);
            for (int index = 1; index < paragraphTexts.Count; index++) {
                string current = paragraphTexts[index];
                int prefixLength = Math.Min(tailLength, current.Length);
                string boundary = tail + current.Substring(0, prefixLength);
                int searchIndex = 0;
                while (searchIndex < tail.Length) {
                    int matchIndex = boundary.IndexOf(searched, searchIndex, stringComparison);
                    if (matchIndex < 0 || matchIndex >= tail.Length) break;
                    if (matchIndex + searched.Length > tail.Length) return true;
                    searchIndex = matchIndex + 1;
                }

                string combined = tail + current;
                tail = combined.Length <= tailLength
                    ? combined
                    : combined.Substring(combined.Length - tailLength);
            }
            return false;
        }

        private static bool HasCrossParagraphOrdinalMatch(
            IReadOnlyList<string> paragraphTexts,
            string searched) {
            int maximumTailLength = searched.Length - 1;
            var tail = new char[maximumTailLength];
            int tailCount = 0;
            for (int index = 0; index < paragraphTexts.Count; index++) {
                string current = paragraphTexts[index];
                if (index > 0 && ContainsOrdinalMatchAcrossBoundary(tail, tailCount, current, searched)) {
                    return true;
                }

                if (current.Length >= maximumTailLength) {
                    current.CopyTo(current.Length - maximumTailLength, tail, 0, maximumTailLength);
                    tailCount = maximumTailLength;
                    continue;
                }

                int retained = Math.Min(tailCount, maximumTailLength - current.Length);
                if (retained > 0) Array.Copy(tail, tailCount - retained, tail, 0, retained);
                current.CopyTo(0, tail, retained, current.Length);
                tailCount = retained + current.Length;
            }
            return false;
        }

        private static bool ContainsOrdinalMatchAcrossBoundary(
            char[] tail,
            int tailCount,
            string current,
            string searched) {
            int combinedLength = tailCount + current.Length;
            int firstStart = Math.Max(0, tailCount - searched.Length + 1);
            for (int start = firstStart; start < tailCount && start + searched.Length <= combinedLength; start++) {
                int matchIndex = 0;
                while (matchIndex < searched.Length) {
                    int combinedIndex = start + matchIndex;
                    char value = combinedIndex < tailCount
                        ? tail[combinedIndex]
                        : current[combinedIndex - tailCount];
                    if (value != searched[matchIndex]) break;
                    matchIndex++;
                }
                if (matchIndex == searched.Length) return true;
            }
            return false;
        }

        private static List<WordTextSegment> SearchText(List<WordParagraph> paragraphs, String searched, WordPositionInParagraph startPos, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {

            var segList = new List<WordTextSegment>();
            int startRun = startPos.Paragraph,
            startText = startPos.Text,
            startChar = startPos.Char;
            int beginRunPos = 0, beginCharPos = 0, candCharPos = 0;
            bool newList = false;
            for (int runPos = startRun; runPos < paragraphs.Count; runPos++) {
                int textPos = 0, charPos = 0;
                var p = paragraphs[runPos];

                if (!string.IsNullOrEmpty(p.Text)) {
                    if (textPos >= startText) {
                        string candidate = p.Text;
                        if (runPos == startRun)
                            charPos = startChar;
                        else
                            charPos = 0;
                        for (; charPos < candidate.Length; charPos++) {
                            if (string.Compare(candidate[charPos].ToString(), searched[0].ToString(), stringComparison) == 0 && (candCharPos == 0)) {
                                beginCharPos = charPos;
                                beginRunPos = runPos;
                                newList = true;
                            }
                            if (string.Compare(candidate[charPos].ToString(), searched[candCharPos].ToString(), stringComparison) == 0) {
                                if (candCharPos + 1 < searched.Length) {
                                    candCharPos++;
                                } else if (newList) {
                                    WordTextSegment segement = new WordTextSegment();
                                    segement.BeginIndex = (beginRunPos);
                                    segement.BeginChar = (beginCharPos);
                                    segement.EndIndex = (runPos);
                                    segement.EndChar = (charPos);
                                    segList.Add(segement);
                                    //Reset
                                    startChar = charPos;
                                    startText = textPos;
                                    startRun = runPos;
                                    newList = false;
                                    candCharPos = 0;
                                }
                            } else
                                candCharPos = 0;
                        }
                    }
                }


            }
            return segList;
        }

        private delegate void SearchScopeVisitor(List<WordParagraph> paragraphs, List<WordParagraph>? destination);

        private void ForEachSearchScope(SearchScopeVisitor visit, WordSearchResult? result) {
            visit(this.Paragraphs, result?.MutableBodyParagraphs);
            foreach (var table in this.Tables) {
                visit(table.Paragraphs, result?.MutableTableParagraphs);
            }

            var headers = this.Sections[0].Headers;
            var footers = this.Sections[0].Footers;
            VisitHeaderFooter(headers.Default, result?.MutableDefaultHeaderParagraphs);
            VisitHeaderFooter(headers.Even, result?.MutableEvenPageHeaderParagraphs);
            VisitHeaderFooter(headers.First, result?.MutableFirstPageHeaderParagraphs);
            VisitHeaderFooter(footers.Default, result?.MutableDefaultFooterParagraphs);
            VisitHeaderFooter(footers.Even, result?.MutableEvenPageFooterParagraphs);
            VisitHeaderFooter(footers.First, result?.MutableFirstPageFooterParagraphs);

            void VisitHeaderFooter(WordHeaderFooter? part, List<WordParagraph>? destination) {
                if (part == null) return;
                visit(part.Paragraphs, destination);
                foreach (var table in part.Tables) {
                    visit(table.Paragraphs, destination);
                }
            }
        }

        private static bool IsSegmentValid(List<WordParagraph> paragraphs, WordTextSegment ts) {
            if (paragraphs is null || ts is null) {
                return false;
            }

            if (ts.BeginIndex < 0 || ts.EndIndex < ts.BeginIndex || ts.EndIndex >= paragraphs.Count) {
                return false;
            }

            var beginPara = paragraphs[ts.BeginIndex];
            var endPara = paragraphs[ts.EndIndex];

            if (beginPara is null || endPara is null) {
                return false;
            }

            if (ts.BeginChar < 0 || ts.BeginChar >= beginPara.Text.Length) {
                return false;
            }

            if (ts.EndChar < 0 || ts.EndChar >= endPara.Text.Length) {
                return false;
            }

            return true;
        }
    }
}
