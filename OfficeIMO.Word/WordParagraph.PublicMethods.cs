using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Collections.Generic;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Add a text to existing paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public WordParagraph AddText(string text) {
            WordParagraph wordParagraph = ConvertToTextWithBreaks(text);
            //WordParagraph wordParagraph = new WordParagraph(this._document, this._paragraph, new Run());
            //wordParagraph.Text = text;
            //this._paragraph.Append(wordParagraph._run);
            return wordParagraph;
        }

        /// <summary>
        /// Add image from file with ability to provide width and height of the image
        /// The image will be resized given new dimensions
        /// </summary>
        /// <param name="filePathImage">Path to file to import to Word Document</param>
        /// <param name="width">Optional width of the image. If not given the actual image width will be used.</param>
        /// <param name="height">Optional height of the image. If not given the actual image height will be used.</param>
        /// <param name="wrapImageText">Optional text wrapping rule. If not given the image will be inserted inline to the text.</param>
        /// <param name="description">The description for this image.</param>
        /// <returns></returns>
        public WordParagraph AddImage(string filePathImage, double? width = null, double? height = null, WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, filePathImage, width, height, wrapImageText, description);
            VerifyRun();
            _run.Append(wordImage._Image);
            return this;
        }
        /// <summary>
        /// Add image from Stream with ability to provide width and height of the image
        /// The image will be resized given new dimensions
        /// </summary>
        /// <param name="imageStream"></param>
        /// <param name="fileName"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="wrapImageText"></param>
        /// <param name="description"></param>
        /// <returns></returns>
        public WordParagraph AddImage(Stream imageStream, string fileName, double? width, double? height, WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, imageStream, fileName, width, height, wrapImageText, description);
            VerifyRun();
            _run.Append(wordImage._Image);
            return this;
        }

        /// <summary>
        /// Add Break to the paragraph. By default it adds soft break (SHIFT+ENTER)
        /// </summary>
        /// <param name="breakType"></param>
        /// <returns></returns>
        public WordParagraph AddBreak(BreakValues? breakType = null) {
            WordParagraph wordParagraph = new WordParagraph(this._document, this._paragraph, new Run());
            if (breakType != null) {
                this._paragraph.Append(new Run(new Break() { Type = breakType }));
            } else {
                this._paragraph.Append(new Run(new Break()));
            }
            return wordParagraph;
        }

        /// <summary>
        /// Remove the paragraph from WordDocument
        /// </summary>
        /// <exception cref="InvalidOperationException"></exception>
        public void Remove() {
            if (_paragraph != null) {
                if (this._paragraph.Parent != null) {
                    if (this.IsBookmark) {
                        this.Bookmark.Remove();
                    }

                    if (this.IsBreak) {
                        this.Break.Remove();
                    }

                    // break should cover this
                    //if (this.IsPageBreak) {
                    //    this.PageBreak.Remove();
                    //}

                    if (this.IsEquation) {
                        this.Equation.Remove();
                    }

                    if (this.IsHyperLink) {
                        this.Hyperlink.Remove();
                    }

                    if (this.IsListItem) {

                    }

                    if (this.IsImage) {
                        this.Image.Remove();
                    }

                    if (this.IsStructuredDocumentTag) {
                        this.StructuredDocumentTag.Remove();
                    }

                    if (this.IsField) {
                        this.Field.Remove();
                    }

                    var runs = this._paragraph.ChildElements.OfType<Run>().ToList();
                    if (runs.Count == 0) {
                        this._paragraph.Remove();
                    } else if (runs.Count == 1) {
                        this._paragraph.Remove();
                    } else {
                        foreach (var run in runs) {
                            if (run == _run) {
                                this._run.Remove();
                            }
                        }
                    }
                } else {
                    throw new InvalidOperationException("This shouldn't happen? Why? Oh why 1?");
                }
            } else {
                // this shouldn't happen
                throw new InvalidOperationException("This shouldn't happen? Why? Oh why 2?");
            }
        }

        /// <summary>
        /// Add paragraph right after existing paragraph.
        /// This can be useful to add empty lines, or moving cursor to next line
        /// </summary>
        /// <param name="wordParagraph"></param>
        /// <returns></returns>
        public WordParagraph AddParagraph(WordParagraph wordParagraph = null) {
            if (wordParagraph == null) {
                // we create paragraph (and within that add it to document)
                wordParagraph = new WordParagraph(this._document, newParagraph: true, newRun: false);
            }
            this._paragraph.InsertAfterSelf(wordParagraph._paragraph);
            return wordParagraph;
        }

        public WordParagraph AddParagraph(string text) {
            // we create paragraph (and within that add it to document)
            var wordParagraph = new WordParagraph(this._document, newParagraph: true, newRun: false) {
                Text = text
            };
            this._paragraph.InsertAfterSelf(wordParagraph._paragraph);
            return wordParagraph;
        }

        /// <summary>
        /// Add paragraph after self adds paragraph after given paragraph
        /// </summary>
        /// <returns></returns>
        public WordParagraph AddParagraphAfterSelf() {
            WordParagraph paragraph = new WordParagraph(this._document, true, false);
            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            return paragraph;
        }

        /// <summary>
        /// Add paragraph after self but by allowing to specify section
        /// </summary>
        /// <param name="section"></param>
        /// <param name="paragraph"></param>
        /// <returns></returns>
        public WordParagraph AddParagraphAfterSelf(WordSection section, WordParagraph paragraph = null) {
            if (paragraph == null) {
                paragraph = new WordParagraph(section._document, true, false);
            }
            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            return paragraph;
        }

        /// <summary>
        /// Add paragraph before another paragraph
        /// </summary>
        /// <returns></returns>
        public WordParagraph AddParagraphBeforeSelf() {
            WordParagraph paragraph = new WordParagraph(this._document, true, false);
            this._paragraph.InsertBeforeSelf(paragraph._paragraph);
            //document.Paragraphs.Add(paragraph);
            return paragraph;
        }

        /// <summary>
        /// Add a comment to paragraph
        /// </summary>
        /// <param name="author"></param>
        /// <param name="initials"></param>
        /// <param name="comment"></param>
        public void AddComment(string author, string initials, string comment) {
            WordComment wordComment = WordComment.Create(_document, author, initials, comment);

            // Specify the text range for the Comment.
            // Insert the new CommentRangeStart before the first run of paragraph.
            this._paragraph.InsertBefore(new CommentRangeStart() { Id = wordComment.Id }, this._paragraph.GetFirstChild<Run>());

            // Insert the new CommentRangeEnd after last run of paragraph.
            var cmtEnd = this._paragraph.InsertAfter(new CommentRangeEnd() { Id = wordComment.Id }, this._paragraph.Elements<Run>().Last());

            // Compose a run with CommentReference and insert it.
            this._paragraph.InsertAfter(new Run(new CommentReference() { Id = wordComment.Id }), cmtEnd);
        }

        /// <summary>
        /// Add horizontal line (sometimes known as horizontal rule) to document
        /// </summary>
        /// <param name="lineType"></param>
        /// <param name="color"></param>
        /// <param name="size"></param>
        /// <param name="space"></param>
        /// <returns></returns>
        public WordParagraph AddHorizontalLine(BorderValues lineType = BorderValues.Single, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            this._paragraphProperties.ParagraphBorders = new ParagraphBorders();
            this._paragraphProperties.ParagraphBorders.BottomBorder = new BottomBorder() {
                Val = lineType,
                Size = size,
                Space = space,
                Color = color != null ? color.Value.ToHexColor() : "auto"
            };
            return this;
        }

        /// <summary>
        /// Add bookmark to a word document
        /// </summary>
        /// <param name="bookmarkName"></param>
        /// <returns></returns>
        public WordParagraph AddBookmark(string bookmarkName) {
            var bookmark = WordBookmark.AddBookmark(this, bookmarkName);
            return this;
        }

        /// <summary>
        /// Add fields to a word document
        /// </summary>
        /// <param name="wordFieldType"></param>
        /// <param name="wordFieldFormat"></param>
        /// <param name="advanced"></param>
        /// <param name="parameters">Usages like <code>parameters = new List&lt; String&gt;{ @"\d 'Default'", @"\c" };</code><br/> Also see available List of switches per field code: <see>https://support.microsoft.com/en-us/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51 </see></param>
        /// <returns></returns>
        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false, List<String> parameters = null) {
            var field = WordField.AddField(this, wordFieldType, wordFieldFormat, advanced, parameters);
            return this;
        }

        /// <summary>
        /// Add hyperlink with URL to a word document
        /// </summary>
        /// <param name="text"></param>
        /// <param name="uri"></param>
        /// <param name="addStyle"></param>
        /// <param name="tooltip"></param>
        /// <param name="history"></param>
        /// <returns></returns>
        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            var hyperlink = WordHyperLink.AddHyperLink(this, text, uri, addStyle, tooltip, history);
            return this;
        }

        /// <summary>
        /// Add hyperlink with an anchor to a word document
        /// </summary>
        /// <param name="text"></param>
        /// <param name="anchor"></param>
        /// <param name="addStyle"></param>
        /// <param name="tooltip"></param>
        /// <param name="history"></param>
        /// <returns></returns>
        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            var hyperlink = WordHyperLink.AddHyperLink(this, text, anchor, addStyle, tooltip, history);
            return this;
        }

        public WordTable AddTableAfter(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this._document, this, rows, columns, tableStyle, "After");
            return wordTable;
        }

        public WordTable AddTableBefore(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this._document, this, rows, columns, tableStyle, "Before");
            return wordTable;
        }

        /// <summary>
        /// Provides ability for configuration of Tabs in a paragraph
        /// by adding one or more TabStops
        /// </summary>
        /// <param name="position"></param>
        /// <param name="alignment"></param>
        /// <param name="leader"></param>
        /// <returns></returns>
        public WordTabStop AddTabStop(int position, TabStopValues alignment = TabStopValues.Left, TabStopLeaderCharValues leader = TabStopLeaderCharValues.None) {
            var wordTab = new WordTabStop(this);
            wordTab.AddTab(position, alignment, leader);
            return wordTab;
        }

        /// <summary>
        /// Adds a Tab to a paragraph
        /// </summary>
        /// <returns></returns>
        public WordParagraph AddTab() {
            var wordParagraph = WordTabChar.AddTab(this._document, this);
            return wordParagraph;
        }

        public WordList AddList(WordListStyle style, bool continueNumbering = false) {
            WordList wordList = new WordList(this._document, this);
            wordList.AddList(style, continueNumbering);
            return wordList;
        }

        public WordChart AddBarChart(string title = null, bool roundedCorners = false, int width = 600, int height = 600) {
            var barChart = WordBarChart.AddBarChart(this._document, this, title, roundedCorners, width, height);
            return barChart;
        }

        public WordChart AddLineChart(string title = null, bool roundedCorners = false, int width = 600, int height = 600) {
            var lineChart = WordLineChart.AddLineChart(this._document, this, title, roundedCorners, width, height);
            return lineChart;
        }

        //public WordBarChart3D AddBarChart3D(string title = null, bool roundedCorners = false, int width = 600, int height = 600) {
        //    var barChart = WordBarChart3D.AddBarChart3D(this._document, this, title, roundedCorners, width, height);
        //    return barChart;
        //}

        public WordChart AddPieChart(string title = null, bool roundedCorners = false, int width = 600, int height = 600) {
            var pieChart = WordPieChart.AddPieChart(this._document, this, title, roundedCorners, width, height);
            return pieChart;
        }

        public WordParagraph AddFootNote(string text) {
            var footerWordParagraph = new WordParagraph(this._document, true, true);
            footerWordParagraph.Text = text;

            var wordFootNote = WordFootNote.AddFootNote(this._document, this, footerWordParagraph);
            return wordFootNote;
        }

        public WordParagraph AddEndNote(string text) {
            var endNoteWordParagraph = new WordParagraph(this._document, true, true);
            endNoteWordParagraph.Text = text;

            var wordEndNote = WordEndNote.AddEndNote(this._document, this, endNoteWordParagraph);
            return wordEndNote;

        }

        public WordTextBox AddTextBox(string text, WrapTextImage wrapTextImage) {
            WordTextBox wordTextBox = new WordTextBox(this._document, this, text, wrapTextImage);
            return wordTextBox;
        }
    }
}
