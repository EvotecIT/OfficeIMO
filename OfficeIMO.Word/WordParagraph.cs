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
    /// <summary>
    /// Represents a paragraph within a Word document.
    /// </summary>
    public partial class WordParagraph : WordElement {
        internal WordDocument _document = null!;
        internal Paragraph _paragraph = null!;
        private object? _parentElement;
        private bool _parentEvaluated;

        /// <summary>
        /// Gets the parent object that owns this paragraph (for example a table cell, header, footer or section).
        /// </summary>
        public object? Parent {
            get {
                if (!_parentEvaluated) {
                    RefreshParent();
                }

                return _parentElement;
            }
            internal set {
                _parentElement = value;
                _parentEvaluated = true;
            }
        }

        internal void RefreshParent() {
            if (_document != null && _paragraph != null) {
                _parentElement = ResolveParent(_document, _paragraph);
            } else {
                _parentElement = null;
            }

            _parentEvaluated = true;
        }

        internal void InvalidateParent() {
            _parentElement = null;
            _parentEvaluated = false;
        }

        /// <summary>
        /// This allows to know where the paragraph is located. Useful for hyperlinks or other stuff.
        /// </summary>
        internal string TopParent {
            get {
                var test = _paragraph.Parent;
                if (test == null) {
                    throw new InvalidOperationException($"Paragraph with text '{Text}' has no parent.");
                }
                if (test is Body) {
                    return "body";
                }
                if (test is Header) {
                    return "header";
                }
                if (test is Footer) {
                    return "footer";
                }
                var parent = test;
                while (!(parent is Header) && !(parent is Footer) && !(parent is Body)) {
                    parent = parent.Parent;
                    if (parent == null) {
                        throw new InvalidOperationException($"Unsupported parent chain for paragraph with text '{Text}'.");
                    }
                }
                if (parent is Body) {
                    return "body";
                }
                if (parent is Footer) {
                    return "footer";
                }
                if (parent is Header) {
                    return "header";
                }
                throw new InvalidOperationException($"Unsupported parent chain for paragraph with text '{Text}'.");
            }
        }

        /// <summary>
        /// Gets a value indicating whether this run is the last run within its parent container.
        /// </summary>
        public bool IsLastRun {
            get {
                if (_run is not null) {
                    var parent = _run.Parent;
                    if (parent != null) {
                        var runs = parent.ChildElements.OfType<Run>();
                        return runs.LastOrDefault() == _run;
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this run is the first run within its parent container.
        /// </summary>
        public bool IsFirstRun {
            get {
                if (_run is not null) {
                    var parent = _run.Parent;
                    if (parent != null) {
                        var runs = parent.ChildElements.OfType<Run>();
                        return runs.FirstOrDefault() == _run;
                    }
                }
                return false;
            }
        }

        internal RunProperties? _runProperties {
            get {
                if (_run is not null) {
                    return _run.RunProperties;
                }

                return null;
            }
            set {
                if (_run != null) {
                    _run.RunProperties = value;
                }
            }
        }

        internal Text? _text {
            get {
                if (_run != null) {
                    return _run.ChildElements.OfType<Text>().FirstOrDefault();
                }

                return null;
            }
        }
        internal Run? _run;

        internal ParagraphProperties? _paragraphProperties {
            get {
                if (_paragraph != null && _paragraph.ParagraphProperties != null) {
                    return _paragraph.ParagraphProperties;
                }

                return null;
            }
        }

        /// <summary>
        /// Gets a value indicating whether this paragraph is part of a numbered or bulleted list.
        /// </summary>
        public bool IsListItem {
            get {
                if (_paragraphProperties != null && _paragraphProperties.NumberingProperties != null) {
                    return true;
                } else {
                    return false;
                }
            }
        }

        /// <summary>
        /// Gets or sets the indentation level for the paragraph when it belongs to a list.
        /// </summary>
        public int? ListItemLevel {
            get {
                var val = _paragraphProperties?.NumberingProperties?.NumberingLevelReference?.Val;
                return val?.Value;
            }
            set {
                var levelRef = _paragraphProperties?.NumberingProperties?.NumberingLevelReference;
                if (levelRef != null) {
                    levelRef.Val = value;
                } else {
                    // should throw?
                }
            }
        }

        internal int? _listNumberId {
            get {
                var val = _paragraphProperties?.NumberingProperties?.NumberingId?.Val;
                return val?.Value;
            }
        }

        /// <summary>
        /// Gets the list style when this paragraph is part of a list.
        /// </summary>
        public WordListStyle? GetListStyle() {
            if (!IsListItem) return null;

            int? numberId = _listNumberId;
            if (numberId == null || _document == null) return null;

            var list = _document.Lists.FirstOrDefault(l => l._numberId == numberId);
            return list?.Style;
        }

        /// <summary>
        /// Gets the list style when this paragraph is part of a list.
        /// </summary>
        public WordListStyle? ListStyle {
            get {
                return GetListStyle();
            }
        }


        /// <summary>
        /// Gets or sets the paragraph style. Updating this to a heading style will flag the document to update the table of contents on open.
        /// </summary>
        public WordParagraphStyles? Style {
            get {
                var styleId = _paragraphProperties?.ParagraphStyleId?.Val;
                return styleId != null ? WordParagraphStyle.GetStyle(styleId.Value!) : null;
            }
            set {
                if (value != null) {
                    if (_paragraphProperties == null) {
                        _paragraph.ParagraphProperties = new ParagraphProperties();
                    }
                    var paragraphProperties = _paragraphProperties;
                    if (paragraphProperties != null) {
                        if (paragraphProperties.ParagraphStyleId == null) {
                            paragraphProperties.ParagraphStyleId = new ParagraphStyleId();
                        }
                        paragraphProperties.ParagraphStyleId.Val = value.Value.ToStringStyle();
                        if (value.Value >= WordParagraphStyles.Heading1 && value.Value <= WordParagraphStyles.Heading9) {
                            _document?.HeadingModified();
                        }
                    }
                }
            }
        }


        internal WordList? _list;
        internal List<Run>? _runs;
        internal Hyperlink? _hyperlink;
        internal SimpleField? _simpleField;
        internal BookmarkStart? _bookmarkStart;
        internal readonly OfficeMath? _officeMath;
        internal readonly SdtRun? _stdRun;
        internal readonly DocumentFormat.OpenXml.Math.Paragraph? _mathParagraph;





        /// <summary>
        /// Initializes a new paragraph.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="newParagraph">Create a new paragraph element.</param>
        /// <param name="newRun">Create a new run inside the paragraph.</param>
        public WordParagraph(WordDocument? document = null, bool newParagraph = true, bool newRun = true) {
            this._document = document!;

            if (newParagraph) {
                this._paragraph = new Paragraph();
                this._paragraph.AppendChild(new ParagraphProperties());

                if (newRun) {
                    this._run = new Run();
                    this._paragraph.AppendChild(_run);
                }
            }

            RefreshParent();
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, bool newRun = true) {
            this._document = document;
            this._paragraph = paragraph;

            if (newRun) {
                this._run = new Run();
                this._paragraph.AppendChild(_run);
            }

            RefreshParent();
        }

        /// <summary>
        /// Wraps an existing paragraph from the document.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph to wrap.</param>
        public WordParagraph(WordDocument document, Paragraph paragraph) {
            this._document = document;
            this._paragraph = paragraph;
            RefreshParent();
        }

        /// <summary>
        /// Wraps an existing paragraph and run from the document.
        /// </summary>
        /// <param name="document">Parent document.</param>
        /// <param name="paragraph">Paragraph to wrap.</param>
        /// <param name="run">Run within the paragraph.</param>
        public WordParagraph(WordDocument document, Paragraph paragraph, Run run) {
            _document = document;
            _paragraph = paragraph;
            _run = run;
            RefreshParent();
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, Hyperlink hyperlink) {
            _document = document;
            _paragraph = paragraph;
            _hyperlink = hyperlink;

            //this.Hyperlink = new WordHyperLink(document, paragraph, hyperlink);
            RefreshParent();
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, List<Run> runs) {
            _document = document;
            _paragraph = paragraph;
            _runs = runs;
            //this.Field = new WordField(document, paragraph, runs);
            RefreshParent();
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, SimpleField simpleField) {
            _document = document;
            _paragraph = paragraph;

            _simpleField = simpleField;

            //  this.Field = new WordField(document, paragraph, simpleField);
            RefreshParent();
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, BookmarkStart bookmarkStart) {
            _document = document;
            _paragraph = paragraph;

            _bookmarkStart = bookmarkStart;

            // this.Bookmark = new WordBookmark(document, paragraph, bookmarkStart);
            RefreshParent();
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
            _document = document;
            _paragraph = paragraph;

            _officeMath = officeMath;

            //this.Equation = new WordEquation(document, paragraph, officeMath);
            RefreshParent();
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, SdtRun stdRun) {
            _document = document;
            _paragraph = paragraph;
            _stdRun = stdRun;
            //this.StructuredDocumentTag = new WordStructuredDocumentTag(document, paragraph, stdRun);
            RefreshParent();
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            _document = document;
            _paragraph = paragraph;
            _mathParagraph = mathParagraph;
            //  this.Equation = new WordEquation(document, paragraph, mathParagraph);
            RefreshParent();
        }

    }
}
