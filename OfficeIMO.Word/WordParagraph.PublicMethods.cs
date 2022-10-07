using System;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Add a text to existing paragraph
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public WordParagraph AddText(string text) {
            WordParagraph wordParagraph = new WordParagraph(this._document, this._paragraph, new Run());
            wordParagraph.Text = text;
            this._paragraph.Append(wordParagraph._run);
            //this._document._wordprocessingDocument.MainDocumentPart.Document.InsertAfter(wordParagraph._run, this._paragraph);
            return wordParagraph;
        }

        public WordParagraph AddImage(string filePathImage, double? width, double? height) {
            WordImage wordImage = new WordImage(this._document, filePathImage, width, height);
            WordParagraph paragraph = new WordParagraph(this._document);
            VerifyRun();
            _run.Append(wordImage._Image);
            //this.Image = wordImage;
            return paragraph;
        }

        public WordParagraph AddImage(string filePathImage) {
            WordImage wordImage = new WordImage(this._document, filePathImage, null, null);
            WordParagraph paragraph = new WordParagraph(this._document);
            VerifyRun();
            _run.Append(wordImage._Image);
            //this.Image = wordImage;
            return paragraph;
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

        public WordParagraph AddParagraphAfterSelf() {
            WordParagraph paragraph = new WordParagraph(this._document, true, false);
            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            return paragraph;
        }

        public WordParagraph AddParagraphAfterSelf(WordSection section) {
            WordParagraph paragraph = new WordParagraph(section._document, true, false);

            this._paragraph.InsertAfterSelf(paragraph._paragraph);

            return paragraph;
        }

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
            //Comments comments = null;
            //string id = "0";

            //// Verify that the document contains a
            //// WordProcessingCommentsPart part; if not, add a new one.
            //if (this._document._wordprocessingDocument.MainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0) {
            //    comments = this._document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments;
            //    if (comments.HasChildren) {
            //        // Obtain an unused ID.
            //        id = (comments.Descendants<Comment>().Select(e => int.Parse(e.Id.Value)).Max() + 1).ToString();
            //    }
            //} else {
            //    // No WordprocessingCommentsPart part exists, so add one to the package.
            //    WordprocessingCommentsPart commentPart = this._document._wordprocessingDocument.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
            //    commentPart.Comments = new Comments();
            //    comments = commentPart.Comments;
            //}

            //// Compose a new Comment and add it to the Comments part.
            //Paragraph p = new Paragraph(new Run(new Text(comment)));
            //Comment cmt =
            //    new Comment() {
            //        Id = id,
            //        Author = author,
            //        Initials = initials,
            //        Date = DateTime.Now
            //    };
            //cmt.AppendChild(p);
            //comments.AppendChild(cmt);
            //comments.Save();

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

        public WordParagraph AddBookmark(string bookmarkName) {
            var bookmark = WordBookmark.AddBookmark(this, bookmarkName);
            return this;
        }

        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, bool advanced = false) {
            var field = WordField.AddField(this, wordFieldType, wordFieldFormat, advanced);
            return this;
        }

        public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle = false, string tooltip = "", bool history = true) {
            var hyperlink = WordHyperLink.AddHyperLink(this, text, uri, addStyle, tooltip, history);
            return this;
        }

        public WordParagraph AddHyperLink(string text, string anchor, bool addStyle = false, string tooltip = "", bool history = true) {
            var hyperlink = WordHyperLink.AddHyperLink(this, text, anchor, addStyle, tooltip, history);
            return this;
        }
    }
}
