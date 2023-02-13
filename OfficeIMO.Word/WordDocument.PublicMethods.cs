using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
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
            return AddParagraph().SetText(text);
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

        public WordParagraph AddBreak(BreakValues breakType = BreakValues.Page) {
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

        public WordChart AddBarChart() {
            var paragraph = this.AddParagraph();
            var barChart = WordBarChart.AddBarChart(this, paragraph);
            return barChart;
        }

        public WordChart AddLineChart() {
            var paragraph = this.AddParagraph();
            var lineChart = WordLineChart.AddLineChart(this, paragraph);
            return lineChart;
        }

        public WordBarChart3D AddBarChart3D() {
            var paragraph = this.AddParagraph();
            var barChart = WordBarChart3D.AddBarChart3D(this, paragraph);
            return barChart;
        }

        public WordChart AddPieChart() {
            var paragraph = this.AddParagraph();
            var pieChart = WordPieChart.AddPieChart(this, paragraph);
            return pieChart;
        }

        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this, this._currentSection);
            wordList.AddList(style);
            return wordList;
        }

        public WordList AddTableOfContentList(WordListStyle style) {
            WordList wordList = new WordList(this, this._currentSection, true);
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

        public WordParagraph AddHorizontalLine(BorderValues lineType = BorderValues.Single, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            return this.AddParagraph().AddHorizontalLine(lineType, color, size, space);
        }

        public WordSection AddSection(SectionMarkValues? sectionMark = null) {
            //Paragraph paragraph = new Paragraph() { RsidParagraphAddition = "fff0", RsidRunAdditionDefault = "fff0"};
            Paragraph paragraph = new Paragraph();

            ParagraphProperties paragraphProperties = new ParagraphProperties();

            SectionProperties sectionProperties = new SectionProperties();
            // SectionProperties sectionProperties = new SectionProperties() { RsidR = "fff0"  };

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

        public WordEmbeddedDocument AddEmbeddedDocument(string fileName, AlternativeFormatImportPartType? type = null) {
            WordEmbeddedDocument embeddedDocument = new WordEmbeddedDocument(this, fileName, type);
            return embeddedDocument;
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

        public int FindAndReplace(string textToFind, string textToReplace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            int countFind = 0;
            FindAndReplaceInternal(textToFind, textToReplace, ref countFind, true, stringComparison);
            return countFind;
        }


        private List<WordParagraph> FindAndReplaceNested(List<WordParagraph> paragraphs, string textToFind, string textToReplace, ref int count, bool replace, StringComparison stringComparison = StringComparison.OrdinalIgnoreCase) {
            List<WordParagraph> foundParagraphs = new List<WordParagraph>();
            foreach (var paragraph in paragraphs) {
                int internalCount = 0;
                var replacedString = paragraph.Text.FindAndReplace(textToFind, textToReplace, stringComparison, ref internalCount);
                if (internalCount > 0) {
                    if (replace) {
                        paragraph.Text = replacedString;
                    }
                    count += internalCount;
                    foundParagraphs.Add(paragraph);
                }
            }

            return foundParagraphs;
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
