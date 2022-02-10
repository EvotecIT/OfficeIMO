using System;
using System.Collections.Generic;
using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Color = DocumentFormat.OpenXml.Wordprocessing.Color;

namespace OfficeIMO.Word {
    public partial class WordParagraph {

        public WordParagraph AddText(string text) {
            WordParagraph wordParagraph = new WordParagraph(this._document, this._paragraph, new Run());
            wordParagraph.Text = text;

            // this ensures that we keep track of matching runs with real paragraphs
            //wordParagraph._linkedParagraph = this;

            //if (this._linkedParagraph != null) {
            //    this._linkedParagraph._paragraph.Append(wordParagraph._run);
            //} else {


            this._paragraph.Append(wordParagraph._run);



            //}

            //this._document._wordprocessingDocument.MainDocumentPart.Document.InsertAfter(wordParagraph._run, this._paragraph);
            return wordParagraph;
        }

        public WordParagraph AddImage(string filePathImage, double? width, double? height) {
            WordImage wordImage = new WordImage(this._document, filePathImage, width, height);
            WordParagraph paragraph = new WordParagraph(this._document);
            _run.Append(wordImage._Image);
            //this.Image = wordImage;
            return paragraph;
        }

        public WordParagraph AddImage(string filePathImage) {
            WordImage wordImage = new WordImage(this._document, filePathImage, null, null);
            WordParagraph paragraph = new WordParagraph(this._document);
            _run.Append(wordImage._Image);
            //this.Image = wordImage;
            return paragraph;
        }

        public void Remove() {
            if (_paragraph != null) {
                if (this._paragraph.Parent != null) {
                    if (this.IsBookmark) {
                        this.Bookmark.Remove();
                    } else {
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
                    }
                } else {
                    throw new InvalidOperationException("This shouldn't happen? Why? Oh why 1?");
                    //Console.WriteLine(this._run);
                }
            } else {
                // this happens if we continue adding to real paragraphs additional runs. In this case we don't need to,
                // delete paragraph, but only remove Runs 
                // this shouldn't happen
                throw new InvalidOperationException("This shouldn't happen? Why? Oh why 2?");
                //this._run.Remove();
            }

            //if (IsPageBreak) {
            //    this._document.PageBreaks.Remove(this);
            //}

            //if (IsListItem) {
            //    if (this._list != null) {
            //        this._list.ListItems.Remove(this);
            //        this._list = null;
            //    }
            //}

            //this._document.Paragraphs.Remove(this);
        }

        public WordParagraph AddParagraphAfterSelf() {
            WordParagraph paragraph = new WordParagraph(this._document, true);
            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            //this._document.Paragraphs.Add(paragraph);

            return paragraph;
        }

        public WordParagraph AddParagraphAfterSelf(WordSection section) {
            //WordParagraph paragraph = new WordParagraph(section._document, true);
            WordParagraph paragraph = new WordParagraph(section._document, true);

            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            //this._document.Paragraphs.Add(paragraph);

            return paragraph;
        }

        public WordParagraph AddParagraphBeforeSelf() {
            WordParagraph paragraph = new WordParagraph(this._document, true);
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
            Comments comments = null;
            string id = "0";

            // Verify that the document contains a 
            // WordProcessingCommentsPart part; if not, add a new one.
            if (this._document._wordprocessingDocument.MainDocumentPart.GetPartsCountOfType<WordprocessingCommentsPart>() > 0) {
                comments = this._document._wordprocessingDocument.MainDocumentPart.WordprocessingCommentsPart.Comments;
                if (comments.HasChildren) {
                    // Obtain an unused ID.
                    id = (comments.Descendants<Comment>().Select(e => int.Parse(e.Id.Value)).Max() + 1).ToString();
                }
            } else {
                // No WordprocessingCommentsPart part exists, so add one to the package.
                WordprocessingCommentsPart commentPart = this._document._wordprocessingDocument.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                commentPart.Comments = new Comments();
                comments = commentPart.Comments;
            }

            // Compose a new Comment and add it to the Comments part.
            Paragraph p = new Paragraph(new Run(new Text(comment)));
            Comment cmt =
                new Comment() {
                    Id = id,
                    Author = author,
                    Initials = initials,
                    Date = DateTime.Now
                };
            cmt.AppendChild(p);
            comments.AppendChild(cmt);
            comments.Save();

            // Specify the text range for the Comment. 
            // Insert the new CommentRangeStart before the first run of paragraph.
            this._paragraph.InsertBefore(new CommentRangeStart() { Id = id }, this._paragraph.GetFirstChild<Run>());

            // Insert the new CommentRangeEnd after last run of paragraph.
            var cmtEnd = this._paragraph.InsertAfter(new CommentRangeEnd() { Id = id }, this._paragraph.Elements<Run>().Last());

            // Compose a run with CommentReference and insert it.
            this._paragraph.InsertAfter(new Run(new CommentReference() { Id = id }), cmtEnd);
        }

        /// <summary>
        /// Add horizontal line (sometimes known as horizontal rule) to document
        /// </summary>
        /// <param name="lineType"></param>
        /// <param name="color"></param>
        /// <param name="size"></param>
        /// <param name="space"></param>
        /// <returns></returns>
        public WordParagraph AddHorizontalLine(BorderValues lineType = BorderValues.Single, System.Drawing.Color? color = null, uint size = 12, uint space = 1) {
            this._paragraphProperties.ParagraphBorders = new ParagraphBorders();
            this._paragraphProperties.ParagraphBorders.BottomBorder = new BottomBorder() {
                Val = lineType,
                Size = size,
                Space = space,
                Color = color != null ? color.Value.ToHexColor() : "auto"
            };

            //newWordParagraph._paragraph = new Paragraph(newWordParagraph._paragraphProperties);

            //this._document._wordprocessingDocument.MainDocumentPart.Document.Body.Append(this._paragraph);
            //this._currentSection.PageBreaks.Add(newWordParagraph);
            //this._currentSection.Paragraphs.Add(newWordParagraph);
            return this;
        }

        public void AddBookmark(string bookmarkName) {
            BookmarkStart bms = new BookmarkStart() { Name = bookmarkName, Id = this._document.BookmarkId.ToString() };
            BookmarkEnd bme = new BookmarkEnd() { Id = this._document.BookmarkId.ToString() };

            var bm = this._run.InsertAfterSelf(bms);
            bm.InsertAfterSelf(bme);
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
