using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using WordDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Hyperlink = DocumentFormat.OpenXml.Wordprocessing.Hyperlink;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using SdtContentPicture = DocumentFormat.OpenXml.Wordprocessing.SdtContentPicture;
using TabStop = DocumentFormat.OpenXml.Wordprocessing.TabStop;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using V = DocumentFormat.OpenXml.Vml;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        /// <summary>
        /// Gets the bookmark associated with this paragraph, if present.
        /// </summary>
        public WordBookmark? Bookmark =>
            _bookmarkStart is not null ? new WordBookmark(_document, _paragraph, _bookmarkStart) : null;

        /// <summary>
        /// Gets the mathematical equation contained in this paragraph, if any.
        /// </summary>
        public WordEquation? Equation {
            get {
                if (_officeMath is not null && _mathParagraph is not null) return new WordEquation(_document, _paragraph, _officeMath, _mathParagraph);
                if (_officeMath is not null) return new WordEquation(_document, _paragraph, _officeMath);
                if (_mathParagraph is not null) return new WordEquation(_document, _paragraph, _mathParagraph);
                if (_simpleField is not null && new WordField(_document, _paragraph, _simpleField, null).FieldType == WordFieldType.EQ) {
                    return new WordEquation(_document, _paragraph, _simpleField);
                }
                if (_runs is not null && new WordField(_document, _paragraph, null, _runs).FieldType == WordFieldType.EQ) {
                    return new WordEquation(_document, _paragraph, _runs);
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the field contained in this paragraph, if any.
        /// </summary>
        public WordField? Field =>
            _simpleField is not null || _runs is not null ? new WordField(_document, _paragraph, _simpleField, _runs) : null;
        /// <summary>
        /// Gets the hyperlink contained in this paragraph, if present.
        /// </summary>
        public WordHyperLink? Hyperlink =>
            _hyperlink is not null ? new WordHyperLink(_document, _paragraph, _hyperlink, _run) : null;

        /// <summary>
        /// Gets the footnote associated with this paragraph, if any.
        /// </summary>
        public WordFootNote? FootNote {
            get {
                if (_run is not null && _runProperties is not null) {
                    var footReference = _run.ChildElements.OfType<FootnoteReference>().FirstOrDefault();
                    if (footReference is not null) {
                        return new WordFootNote(_document, _paragraph, _run);
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets the endnote associated with this paragraph, if any.
        /// </summary>
        public WordEndNote? EndNote {
            get {
                if (_run is not null && _runProperties is not null) {
                    var endNoteReference = _run.ChildElements.OfType<EndnoteReference>().FirstOrDefault();
                    if (endNoteReference is not null) {
                        return new WordEndNote(_document, _paragraph, _run);
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the paragraph contains a hyperlink.
        /// </summary>
        public bool IsHyperLink => Hyperlink is not null;

        /// <summary>
        /// Gets a value indicating whether the paragraph hosts a field code.
        /// </summary>
        public bool IsField {
            get {
                var wf = Field;
                return wf is not null && wf.Field is not null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the paragraph includes a bookmark start tag.
        /// </summary>
        public bool IsBookmark {
            get {
                var bookmark = Bookmark;
                return bookmark is not null && bookmark.Name is not null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the paragraph contains an equation.
        /// </summary>
        public bool IsEquation => Equation is not null;
        /// <summary>
        /// Gets a value indicating whether an endnote reference is present in the paragraph.
        /// </summary>
        public bool IsEndNote => EndNote is not null;

        /// <summary>
        /// Gets a value indicating whether a footnote reference is present in the paragraph.
        /// </summary>
        public bool IsFootNote => FootNote is not null;
    }
}
