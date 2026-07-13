using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Reflection;
using System.Linq;
using System.Xml.Linq;
using MathParagraph = DocumentFormat.OpenXml.Math.Paragraph;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using V = DocumentFormat.OpenXml.Vml;
using W14 = DocumentFormat.OpenXml.Office2010.Word;

namespace OfficeIMO.Word {
    /// <summary>
    /// Contains public methods for editing paragraphs.
    /// </summary>
    public partial class WordParagraph {
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
        public WordParagraph AddHorizontalLine(BorderValues? lineType = null, OfficeIMO.Drawing.OfficeColor? color = null, uint size = 12, uint space = 1) {
            lineType ??= BorderValues.Single;
            var paragraphProperties = _paragraph!.ParagraphProperties ??= new ParagraphProperties();
            paragraphProperties.ParagraphBorders = new ParagraphBorders {
                BottomBorder = new BottomBorder() {
                    Val = lineType.Value,
                    Size = size,
                    Space = space,
                    Color = color != null ? color.Value.ToRgbHex() : "auto"
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
        /// Adds a cross reference field to the paragraph.
        /// </summary>
        /// <param name="referenceId">Bookmark or heading identifier.</param>
        /// <param name="type">Type of reference.</param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddCrossReference(string referenceId, WordCrossReferenceType type) {
            WordCrossReference.AddCrossReference(this, referenceId, type);
            return this;
        }

        /// <summary>
        /// Adds a citation field referencing the specified source tag.
        /// </summary>
        /// <param name="sourceTag">Tag of the bibliographic source.</param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddCitation(string sourceTag) {
            var field = new CitationField { SourceTag = sourceTag };
            WordField.AddField(this, field, null!, null!, false);
            return this;
        }

        /// <summary>
        /// Add fields to a word document proceeding from the paragraph this is called on.
        /// </summary>
        /// <param name="wordFieldType">The type of field to add.</param>
        /// <param name="wordFieldFormat">The format of the field to add.</param>
        /// <param name="customFormat">Custom format string for date or time fields.</param>
        /// <param name="advanced">Use advanced field representation.</param>
        /// <param name="parameters">Usages like <code>parameters = new List&lt; String&gt;{ @"\d 'Default'", @"\c" };</code><br/>
        /// Also see available List of switches per field code:
        /// <see>https://support.microsoft.com/en-us/office/list-of-field-codes-in-word-1ad6d91a-55a7-4a8d-b535-cf7888659a51 </see></param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddField(WordFieldType wordFieldType, WordFieldFormat? wordFieldFormat = null, string? customFormat = null, bool advanced = false, List<string>? parameters = null) {
            var field = WordField.AddField(this, wordFieldType, wordFieldFormat, customFormat!, advanced, parameters!);
            return this;
        }

        /// <summary>
        /// Add a field represented by a <see cref="WordFieldCode"/>.
        /// </summary>
        /// <param name="fieldCode">Field code instance describing instructions and switches.</param>
        /// <param name="wordFieldFormat">Optional field format.</param>
        /// <param name="customFormat">Custom format string for date or time fields.</param>
        /// <param name="advanced">Use advanced field representation.</param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddField(WordFieldCode fieldCode, WordFieldFormat? wordFieldFormat = null, string? customFormat = null, bool advanced = false) {
            WordField.AddField(this, fieldCode, wordFieldFormat, customFormat!, advanced);
            return this;
        }

        /// <summary>
        /// Adds a field constructed using <see cref="WordFieldBuilder"/>.
        /// </summary>
        /// <param name="builder">Field builder instance.</param>
        /// <param name="advanced">Use advanced field representation.</param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddField(WordFieldBuilder builder, bool advanced = false) {
            WordField.AddField(this, builder, advanced);
            return this;
        }

        /// <summary>
        /// Adds a page number field to the paragraph.
        /// </summary>
        /// <param name="includeTotalPages">If true adds a NUMPAGES field preceded by text " of ".</param>
        /// <param name="format">Optional field format to apply.</param>
        /// <param name="separator">Text inserted between the current page and total pages fields.</param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddPageNumber(bool includeTotalPages = false, WordFieldFormat? format = null, string separator = " of ") {
            this.AddField(WordFieldType.Page, format);
            if (includeTotalPages) {
                this.AddText(separator);
                this.AddField(WordFieldType.NumPages, format);
            }
            return this;
        }

        /// <summary>
        /// Adds a mathematical equation represented as OMML XML.
        /// </summary>
        /// <param name="omml">Office Math Markup Language (OMML) fragment.</param>
        /// <returns>The paragraph that this was called on.</returns>
        public WordParagraph AddEquation(string omml) {
            if (string.IsNullOrWhiteSpace(omml)) {
                throw new ArgumentNullException(nameof(omml));
            }

            XElement x = XElement.Parse(omml);
            WordParagraph paragraphWithEquation;

            if (x.Name.LocalName == "oMath") {
                var officeMath = new OfficeMath(omml);
                _paragraph.Append(officeMath);
                paragraphWithEquation = new WordParagraph(this._document, this._paragraph, officeMath);
            } else {
                var mathPara = new MathParagraph(omml);
                _paragraph.Append(mathPara);
                paragraphWithEquation = new WordParagraph(this._document, this._paragraph, mathPara);
            }

            return paragraphWithEquation;
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
        /// Removes hyperlink from this paragraph and detaches its relationship.
        /// </summary>
        /// <param name="includingParagraph">If true removes the paragraph when it becomes empty.</param>
        public void RemoveHyperLink(bool includingParagraph = false) {
            if (_hyperlink != null) {
                if (!string.IsNullOrEmpty(_hyperlink.Id)) {
                    OpenXmlElement? parent = _paragraph.Parent;
                    while (parent != null && parent is not Body and not Header and not Footer) {
                        parent = parent.Parent;
                    }

                    OpenXmlPart? part = _document._wordprocessingDocument.MainDocumentPart;
                    if (parent is Header header) {
                        part = header.HeaderPart;
                    } else if (parent is Footer footer) {
                        part = footer.FooterPart;
                    }

                    var rel = part?.HyperlinkRelationships.FirstOrDefault(r => r.Id == _hyperlink.Id);
                    if (rel != null && part != null) {
                        part.DeleteReferenceRelationship(rel);
                    }
                }

                _hyperlink.Remove();
                _hyperlink = null!;

                if (includingParagraph) {
                    if (this._paragraph.ChildElements.Count == 0) {
                        this._paragraph.Remove();
                    } else if (this._paragraph.ChildElements.Count == 1 && this._paragraph.ChildElements.OfType<ParagraphProperties>().Any()) {
                        this._paragraph.Remove();
                    }
                }
            }
        }
    }
}