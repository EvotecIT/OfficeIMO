using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        public WordParagraph AddParagraph(WordParagraph wordParagraph = null) {
            if (wordParagraph == null) {
                // we create paragraph (and within that add it to document)
                wordParagraph = new WordParagraph(this, newParagraph: true, newRun: false);
            }

            this._wordprocessingDocument.MainDocumentPart.Document.Body.AppendChild(wordParagraph._paragraph);
            return wordParagraph;
        }

        public WordParagraph AddParagraph(string text) {
            //return AddParagraph().SetText(text);
            return AddParagraph().AddText(text);
        }

        public WordParagraph AddPageBreak() {
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() { Type = BreakValues.Page }),
                _document = this
            };
            newWordParagraph._paragraph = new Paragraph(newWordParagraph._run);

            this._document.Body.Append(newWordParagraph._paragraph);
            return newWordParagraph;
        }

        public void AddHeadersAndFooters() {
            WordHeadersAndFooters.AddHeadersAndFooters(this);
        }

        public WordParagraph AddBreak(BreakValues? breakType = null) {
            breakType ??= BreakValues.Page;
            WordParagraph newWordParagraph = new WordParagraph {
                _run = new Run(new Break() { Type = breakType }),
                _document = this
            };
            newWordParagraph._paragraph = new Paragraph(newWordParagraph._run);

            this._document.Body.Append(newWordParagraph._paragraph);
            this.Paragraphs.Add(newWordParagraph);
            return newWordParagraph;
        }

        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, uri, addStyle, tooltip, history);
        }

        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            return this.AddParagraph().AddHyperLink(text, anchor, addStyle, tooltip, history);
        }

        /// <summary>
        /// Adds the chart to the document. The type of chart is determined by the type of data passed in.
        /// You can use multiple:
        /// .AddBar() to add a bar chart
        /// .AddLine() to add a line chart
        /// .AddPie() to add a pie chart
        /// .AddArea() to add an area chart.
        /// You can't mix and match the types of charts.
        /// </summary>
        /// <param name="title">The title.</param>
        /// <param name="roundedCorners">if set to <c>true</c> [rounded corners].</param>
        /// <param name="width">The width.</param>
        /// <param name="height">The height.</param>
        /// <returns>WordChart</returns>
        public WordChart AddChart(string title = "", bool roundedCorners = false, int width = 600, int height = 600) {
            var paragraph = this.AddParagraph();
            var chartInstance = new WordChart(this, paragraph, title, roundedCorners, width, height);
            return chartInstance;
        }

        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this);
            wordList.AddList(style);
            return wordList;
        }

        public WordList AddTableOfContentList(WordListStyle style) {
            WordList wordList = new WordList(this, true);
            wordList.AddList(style);
            return wordList;
        }

        public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this, rows, columns, tableStyle);
            return wordTable;
        }

        public WordTableOfContent AddTableOfContent(TableOfContentStyle tableOfContentStyle = TableOfContentStyle.Template1) {
            WordTableOfContent wordTableContent = new WordTableOfContent(this, tableOfContentStyle);
            return wordTableContent;
        }

        public WordCoverPage AddCoverPage(CoverPageTemplate coverPageTemplate) {
            WordCoverPage wordCoverPage = new WordCoverPage(this, coverPageTemplate);
            return wordCoverPage;
        }

        public WordTextBox AddTextBox(string text, WrapTextImage wrapTextImage = WrapTextImage.Square) {
            WordTextBox wordTextBox = new WordTextBox(this, text, wrapTextImage);
            return wordTextBox;
        }

        public WordParagraph AddHorizontalLine(BorderValues? lineType = null, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            lineType ??= BorderValues.Single;
            return this.AddParagraph().AddHorizontalLine(lineType.Value, color, size, space);
        }

        public WordSection AddSection(SectionMarkValues? sectionMark = null) {
            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProperties = new ParagraphProperties();

            SectionProperties sectionProperties = WordHeadersAndFooters.CreateSectionProperties();

            if (sectionMark != null) {
                SectionType sectionType = new SectionType() { Val = sectionMark };
                sectionProperties.Append(sectionType);
            }

            paragraphProperties.Append(sectionProperties);
            paragraph.Append(paragraphProperties);


            this._document.MainDocumentPart.Document.Body.Append(paragraph);


            WordSection wordSection = new WordSection(this, paragraph);

            return wordSection;
        }

        public WordParagraph AddBookmark(string bookmarkName) {
            return this.AddParagraph().AddBookmark(bookmarkName);
        }

        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false, List<String> parameters = null) {
            return this.AddParagraph().AddField(wordFieldType, wordFieldFormat, advanced, parameters);
        }

        public WordEmbeddedDocument AddEmbeddedDocument(string fileName, WordAlternativeFormatImportPartType? type = null) {
            return new WordEmbeddedDocument(this, fileName, type, false);
        }

        public WordEmbeddedDocument AddEmbeddedFragment(string htmlContent, WordAlternativeFormatImportPartType type) {
            return new WordEmbeddedDocument(this, htmlContent, type, true);
        }


        private int CombineRuns(WordHeaderFooter wordHeaderFooter) {
            int count = 0;
            if (wordHeaderFooter != null) {
                foreach (var p in this.Header.Default.Paragraphs) count += CombineIdenticalRuns(p._paragraph);
                foreach (var table in this.Header.Default.Tables) {
                    table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
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
        public int CleanupDocument() {
            int count = 0;

            foreach (var paragraph in this.Paragraphs) {
                count += CombineIdenticalRuns(paragraph._paragraph);
            }

            foreach (var table in this.Tables) {
                table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
            }

            if (this.Header.Default != null) {
                foreach (var p in this.Header.Default.Paragraphs) count += CombineIdenticalRuns(p._paragraph);
                foreach (var table in this.Header.Default.Tables) {
                    table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                }
            }

            if (this.Header.Even != null) {
                this.Header.Even.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                foreach (var table in this.Header.Even.Tables) {
                    table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                }
            }
            if (this.Header.First != null) {
                this.Header.First.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                foreach (var table in this.Header.First.Tables) {
                    table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                }
            }

            if (this.Footer.Default != null) {
                this.Footer.Default.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                foreach (var table in this.Footer.Default.Tables) {
                    table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                }
            }

            if (this.Footer.Even != null) {
                this.Footer.Even.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                foreach (var table in this.Footer.Even.Tables) {
                    table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                }
            }

            if (this.Footer.First != null) {
                this.Footer.First.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                foreach (var table in this.Footer.First.Tables) {
                    table.Paragraphs.ForEach(p => count += CombineIdenticalRuns(p._paragraph));
                }
            }
            return count;
        }

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
            var foundList = SearchText(paragraphs, oldText, new WordPositionInParagraph() { Paragraph = 0 });

            if (foundList?.Count > 0) {
                count += foundList.Count;
                foreach (var ts in foundList) {
                    if (ts == null)
                        continue;
                    if (ts.BeginIndex == ts.EndIndex) {
                        var p = paragraphs[ts.BeginIndex];
                        if (p != null) {
                            if (replace) {
                                p.Text = p.Text.Replace(oldText, newText);
                            }
                            if (foundParagraphs.IndexOf(p) == -1) {
                                foundParagraphs.Add(p);
                            }
                        }
                    } else {
                        if (replace) {
                            var beginPara = paragraphs[ts.BeginIndex];
                            var endPara = paragraphs[ts.EndIndex];
                            if (beginPara != null && endPara != null) {
                                beginPara.Text = beginPara.Text.Replace(beginPara.Text.Substring(ts.BeginChar), newText);
                                endPara.Text = endPara.Text.Replace(endPara.Text.Substring(0, ts.EndChar + 1), "");
                                if (foundParagraphs.IndexOf(beginPara) == -1) {
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
                    textPos++;
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

            if (this.Header.Default != null) {
                list.AddRange(FindAndReplaceNested(this.Header.Default.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Header.Default.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Header.Even != null) {
                list.AddRange(FindAndReplaceNested(this.Header.Even.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Header.Even.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Header.First != null) {
                list.AddRange(FindAndReplaceNested(this.Header.First.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Header.First.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Footer.Default != null) {
                list.AddRange(FindAndReplaceNested(this.Footer.Default.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Footer.Default.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Footer.Even != null) {
                list.AddRange(FindAndReplaceNested(this.Footer.Even.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Footer.Even.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            if (this.Footer.First != null) {
                list.AddRange(FindAndReplaceNested(this.Footer.First.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                foreach (var table in this.Footer.First.Tables) {
                    list.AddRange(FindAndReplaceNested(table.Paragraphs, textToFind, textToReplace, ref count, replace, stringComparison));
                }
            }

            return list;
        }
    }
}
