using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Http;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private int CombineRuns(WordHeaderFooter wordHeaderFooter) {
            int count = 0;
            if (wordHeaderFooter != null) {
                var defaultHeader = this.Header?.Default;
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
        public int CleanupDocument(DocumentCleanupOptions options = DocumentCleanupOptions.All) {
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
        /// Searches the document for paragraphs containing the specified text.
        /// </summary>
        /// <param name="text">Text to search for.</param>
        /// <param name="stringComparison">Comparison rules for the search.</param>
        /// <returns>A list of found <see cref="WordParagraph"/> instances.</returns>
        public List<WordParagraph> Find(string text, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            int count = 0;
            List<WordParagraph> list = FindAndReplaceInternal(text, "", ref count, false, stringComparison);
            return list;
        }

        /// <summary>
        /// FindAdnReplace from the whole doc
        /// </summary>
        /// <param name="textToFind"></param>
        /// <param name="textToReplace"></param>
        /// <param name="stringComparison"></param>
        /// <returns></returns>
        public int FindAndReplace(string textToFind, string textToReplace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            int countFind = 0;
            FindAndReplaceInternal(textToFind, textToReplace, ref countFind, true, stringComparison);
            return countFind;
        }

        /// <summary>
        /// FindAdnReplace from the range parparagraphs
        /// </summary>
        /// <param name="paragraphs"></param>
        /// <param name="textToFind"></param>
        /// <param name="textToReplace"></param>
        /// <param name="stringComparison"></param>
        /// <returns></returns>
        public static int FindAndReplace(List<WordParagraph> paragraphs, string textToFind, string textToReplace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            int countFind = 0;
            FindAndReplaceNested(paragraphs, textToFind, textToReplace, ref countFind, true, stringComparison);
            return countFind;
        }


        private static List<WordParagraph> FindAndReplaceNested(List<WordParagraph> paragraphs, string textToFind, string textToReplace, ref int count, bool replace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            List<WordParagraph> foundParagraphs = ReplaceText(paragraphs, textToFind, textToReplace, ref count, replace, stringComparison);
            return foundParagraphs;
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
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        private static List<WordParagraph> ReplaceText(List<WordParagraph> paragraphs, string oldText, string newText, ref int count, bool replace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            if (string.IsNullOrEmpty(oldText)) {
                throw new ArgumentNullException("oldText should not be null");
            }
            List<WordParagraph> foundParagraphs = new List<WordParagraph>();
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

        private List<WordParagraph> FindAndReplaceInternal(string textToFind, string textToReplace, ref int count, bool replace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            WordFind wordFind = new WordFind();
            List<WordParagraph> list = new List<WordParagraph>();
            list.AddRange(FindAndReplaceNested(this.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));

            foreach (var table in this.Tables) {
                list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
            }

            if (this.Header?.Default != null) {
                list.AddRange(FindAndReplaceNested(this.Header.Default.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Header.Default.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Header?.Even != null) {
                list.AddRange(FindAndReplaceNested(this.Header.Even.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Header.Even.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Header?.First != null) {
                list.AddRange(FindAndReplaceNested(this.Header.First.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Header.First.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Footer?.Default != null) {
                list.AddRange(FindAndReplaceNested(this.Footer.Default.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Footer.Default.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Footer?.Even != null) {
                list.AddRange(FindAndReplaceNested(this.Footer.Even.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Footer.Even.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Footer?.First != null) {
                list.AddRange(FindAndReplaceNested(this.Footer.First.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Footer.First.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            return list;
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
