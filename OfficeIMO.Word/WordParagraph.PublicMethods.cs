using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using System.Collections.Generic;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        // Should the return type be changed to signify the difference between paragraph and run?
        /// <summary>
        /// Add a text to existing paragraph
        /// </summary>
        /// <param name="text">The text to be added to the paragraph.</param>
        /// <returns>The paragraph containing the new text.</returns>
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
        /// <returns>The WordParagraph that AddImage was called on.</returns>
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
        /// <param name="imageStream">The stream to load the image from.</param>
        /// <param name="fileName">The name of the file.</param>
        /// <param name="width">Optional width of the image. If not given the actual image width will be used.</param>
        /// <param name="height">Optional height of the image. If not give the actual image height will be used.</param>
        /// <param name="wrapImageText">Optional text wrapping rule. If not given the image will be inserted inline to the text.</param>
        /// <param name="description">The description for this image.</param>
        /// <returns>The WordParagraph that AddImage was called on.</returns>
        public WordParagraph AddImage(Stream imageStream, string fileName, double? width, double? height, WrapTextImage wrapImageText = WrapTextImage.InLineWithText, string description = "") {
            var wordImage = new WordImage(_document, this, imageStream, fileName, width, height, wrapImageText, description);
            VerifyRun();
            _run.Append(wordImage._Image);
            return this;
        }

        /// <summary>
        /// Add Break to the paragraph. By default it adds soft break (SHIFT+ENTER)
        /// </summary>
        /// <param name="breakType">Optional argument to add a specific type of break.</param>
        /// <returns>The new WordParagraph that this method creates.</returns>
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
        /// <param name="wordParagraph">Optional WordParagraph to insert after the
        /// given WordParagraph instead of at the end of the document.</param>
        /// <returns>The inserted WordParagraph.</returns>
        public WordParagraph AddParagraph(WordParagraph wordParagraph = null) {
            if (wordParagraph == null) {
                // we create paragraph (and within that add it to document)
                wordParagraph = new WordParagraph(this._document, newParagraph: true, newRun: false);
            }
            this._paragraph.InsertAfterSelf(wordParagraph._paragraph);
            return wordParagraph;
        }

        /// <summary>
        /// Add a paragraph with the given text to the end of the document.
        /// </summary>
        /// <param name="text">The text to fill the paragraph with.</param>
        /// <returns> The appended WordParagraph.</returns>
        public WordParagraph AddParagraph(string text) {
            // we create paragraph (and within that add it to document)
            var wordParagraph = new WordParagraph(this._document, newParagraph: true, newRun: false) {
                Text = text
            };
            this._paragraph.InsertAfterSelf(wordParagraph._paragraph);
            return wordParagraph;
        }

        /// <summary>
        /// Insert a new paragraph after the WordParagraph AddParagraphAfterSelf is called on in the document.
        /// </summary>
        /// <returns>The inserted WordParagraph</returns>
        public WordParagraph AddParagraphAfterSelf() {
            WordParagraph paragraph = new WordParagraph(this._document, true, false);
            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            return paragraph;
        }

        /// <summary>
        /// Add paragraph after self but by allowing to specify section
        /// </summary>
        /// <param name="section">The section to add the paragraph to. When paragraph is given this has no effect.</param>
        /// <param name="paragraph">The optional paragraph to add the paragraph to.</param>
        /// <returns>The new WordParagraph</returns>
        public WordParagraph AddParagraphAfterSelf(WordSection section, WordParagraph paragraph = null) {
            if (paragraph == null) {
                paragraph = new WordParagraph(section._document, true, false);
            }
            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            return paragraph;
        }

        /// <summary>
        /// Add a paragraph before the paragraph that AddParagraphBeforeSelf was called on.
        /// </summary>
        /// <returns>The inserted paragraph</returns>
        public WordParagraph AddParagraphBeforeSelf() {
            WordParagraph paragraph = new WordParagraph(this._document, true, false);
            this._paragraph.InsertBeforeSelf(paragraph._paragraph);
            //document.Paragraphs.Add(paragraph);
            return paragraph;
        }

        // Should author and initials be made optional or should the user handle that with ""?
        /// <summary>
        /// Add a comment to paragraph
        /// </summary>
        /// <param name="author">The name of the commenting author</param>
        /// <param name="initials">The initials of the commenting author</param>
        /// <param name="comment">The comment text</param>
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

        // Does this return the paragraph after the line, or does this return the paragraph containing the line?
        /// <summary>
        /// Add horizontal line (sometimes known as horizontal rule) to document proceeding from the paragraph that this is called on.
        /// </summary>
        /// <param name="lineType">The type of the line.</param>
        /// <param name="color">The color of the line</param>
        /// <param name="size">The size of the line.</param>
        /// <param name="space">The space the line takes up.</param>
        /// <returns>The new Paragraph after the line.</returns>
        public WordParagraph AddHorizontalLine(BorderValues? lineType = null, SixLabors.ImageSharp.Color? color = null, uint size = 12, uint space = 1) {
            lineType ??= BorderValues.Single;
            this._paragraphProperties.ParagraphBorders = new ParagraphBorders {
                BottomBorder = new BottomBorder() {
                    Val = lineType.Value,
                    Size = size,
                    Space = space,
                    Color = color != null ? color.Value.ToHexColor() : "auto"
                }
            };
            return this;
        }

        /// <summary>
        /// Add bookmark to a word document proceeding from the paragraph this was called on.
        /// </summary>
        /// <param name="bookmarkName">The name of the bookmark.</param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddBookmark(string bookmarkName) {
            var bookmark = WordBookmark.AddBookmark(this, bookmarkName);
            return this;
        }

        /// <summary>
        /// Add fields to a word document proceeding from the paragraph this is called on.
        /// </summary>
        /// <param name="wordFieldType">The type of field to add.</param>
        /// <param name="wordFieldFormat">The format of the field to add.</param>
        /// <param name="advanced"></param>
        /// <param name="parameters">Usages like <code>parameters = new List&lt; String&gt;{ @"\d 'Default'", @"\c" };</code><br/>
        /// Also see available List of switches per field code:
        /// <see>https://support.microsoft.com/en-us/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51 </see></param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false, List<String> parameters = null) {
            var field = WordField.AddField(this, wordFieldType, wordFieldFormat, advanced, parameters);
            return this;
        }

        /// <summary>
        /// Add hyperlink with URL to a word document proceding from the paragraph that this was called on.
        /// </summary>
        /// <param name="text">The text to insert as the URL.</param>
        /// <param name="uri">The uri that this points to.</param>
        /// <param name="addStyle">The optional style of the link.</param>
        /// <param name="tooltip">The optional tooltip to display over the link.</param>
        /// <param name="history"></param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            var hyperlink = WordHyperLink.AddHyperLink(this, text, uri, addStyle, tooltip, history);
            return this;
        }

        /// <summary>
        /// Add hyperlink with an anchor to a word document proceding from the paragraph that this was called on.
        /// </summary>
        /// <param name="text">The text to insert as the URL.</param>
        /// <param name="anchor">The anchor to point at.</param>
        /// <param name="addStyle">The optional style of this link.</param>
        /// <param name="tooltip">The optional tooltip over this link.</param>
        /// <param name="history"></param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            var hyperlink = WordHyperLink.AddHyperLink(this, text, anchor, addStyle, tooltip, history);
            return this;
        }

        /// <summary>
        /// Add a table after this paragraph and return the table.
        /// </summary>
        /// <param name="rows">The number of rows in the table.</param>
        /// <param name="columns">The number of columns in the table.</param>
        /// <param name="tableStyle">The optional style to be applied to the new table, defaults to TableGrid.</param>
        /// <returns>The new added table.</returns>
        public WordTable AddTableAfter(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this._document, this, rows, columns, tableStyle, "After");
            return wordTable;
        }

        /// <summary>
        /// Add a table before this paragraph and return the table.
        /// </summary>
        /// <param name="rows">The number of rows in this table</param>
        /// <param name="columns">The number of columns in this table</param>
        /// <param name="tableStyle">The style of the table being added.</param>
        /// <returns>The new added table.</returns>
        public WordTable AddTableBefore(int rows, int columns, WordTableStyle tableStyle = WordTableStyle.TableGrid) {
            WordTable wordTable = new WordTable(this._document, this, rows, columns, tableStyle, "Before");
            return wordTable;
        }

        /// <summary>
        /// Provides ability for configuration of Tabs in a paragraph
        /// by adding one or more TabStops
        /// </summary>
        /// <param name="position">The position of the tabs in the paragraph.</param>
        /// <param name="alignment">The optional alignment for the tabs.</param>
        /// <param name="leader">The optional rune to use before the tabs.</param>
        /// <returns>The added tabs.</returns>
        public WordTabStop AddTabStop(int position, TabStopValues? alignment = null, TabStopLeaderCharValues? leader = null) {
            alignment ??= TabStopValues.Left;
            leader ??= TabStopLeaderCharValues.None;
            var wordTab = new WordTabStop(this);
            wordTab.AddTab(position, alignment, leader);
            return wordTab;
        }

        /// <summary>
        /// Adds a Tab to a paragraph
        /// </summary>
        /// <returns>The paragraph this is called on.</returns>
        public WordParagraph AddTab() {
            var wordParagraph = WordTabChar.AddTab(this._document, this);
            return wordParagraph;
        }

        /// <summary>
        /// Add a list after this paragraph.
        /// </summary>
        /// <param name="style">The style of this list.</param>
        /// <returns>The new list.</returns>
        public WordList AddList(WordListStyle style) {
            WordList wordList = new WordList(this._document, this);
            wordList.AddList(style);
            return wordList;
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
            var chartInstance = new WordChart(this._document, paragraph, title, roundedCorners, width, height);
            return chartInstance;
        }

        /// <summary>
        /// Add a foot note for the current paragraph.
        /// </summary>
        /// <param name="text">The text of the note.</param>
        /// <returns>The footnote.</returns>
        public WordParagraph AddFootNote(string text) {
            var footerWordParagraph = new WordParagraph(this._document, true, true) {
                Text = text
            };

            var wordFootNote = WordFootNote.AddFootNote(this._document, this, footerWordParagraph);
            return wordFootNote;
        }

        /// <summary>
        /// Add an end note to the document for this paragraph.
        /// </summary>
        /// <param name="text">The text of the end note.</param>
        /// <returns>The end note.</returns>
        public WordParagraph AddEndNote(string text) {
            var endNoteWordParagraph = new WordParagraph(this._document, true, true);
            endNoteWordParagraph.Text = text;

            var wordEndNote = WordEndNote.AddEndNote(this._document, this, endNoteWordParagraph);
            return wordEndNote;

        }

        /// <summary>
        /// Add a text box to the document.
        /// </summary>
        /// <param name="text">The text inside the text box.</param>
        /// <param name="wrapTextImage">The text image wrapping settings.</param>
        public WordTextBox AddTextBox(string text, WrapTextImage wrapTextImage) {
            WordTextBox wordTextBox = new WordTextBox(this._document, this, text, wrapTextImage);
            return wordTextBox;
        }
    }
}
