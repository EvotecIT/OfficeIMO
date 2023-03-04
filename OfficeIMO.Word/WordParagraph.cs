using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using OfficeMath = DocumentFormat.OpenXml.Math.OfficeMath;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        internal WordDocument _document;
        internal Paragraph _paragraph;

        /// <summary>
        /// This allows to know where the paragraph is located. Useful for hyperlinks or other stuff.
        /// </summary>
        internal string Parent {
            get {
                var test = _paragraph.Parent;
                if (test is Body) {
                    return "body";
                } else if (test is Header) {
                    return "header";
                } else if (test is Footer) {
                    return "footer";
                } else {
                    throw new NotImplementedException("There is different parent for paragraphs?");
                }
            }
        }

        public bool IsLastRun {
            get {
                var runs = _run.Parent.ChildElements.OfType<Run>();
                return runs.Last() == _run;
            }
        }

        public bool IsFirstRun {
            get {
                var runs = _run.Parent.ChildElements.OfType<Run>();
                return runs.First() == _run;
            }
        }

        internal RunProperties _runProperties {
            get {
                if (_run != null) {
                    return _run.RunProperties;
                }

                return null;
            }
        }

        internal Text _text {
            get {
                if (_run != null) {
                    return _run.ChildElements.OfType<Text>().FirstOrDefault();
                }

                return null;
            }
        }
        internal Run _run;

        internal ParagraphProperties _paragraphProperties {
            get {
                if (_paragraph != null && _paragraph.ParagraphProperties != null) {
                    return _paragraph.ParagraphProperties;
                }

                return null;
            }
        }
        //internal WordParagraph _linkedParagraph;
        //internal WordSection _section;

        public WordImage Image {
            get {
                if (_run != null) {
                    var drawing = _run.ChildElements.OfType<Drawing>().FirstOrDefault();
                    if (drawing != null) {
                        return new WordImage(_document, drawing);
                    }
                }

                return null;
            }
        }

        public bool IsListItem {
            get {
                if (_paragraphProperties != null && _paragraphProperties.NumberingProperties != null) {
                    return true;
                } else {
                    return false;
                }
            }
        }

        public int? ListItemLevel {
            get {
                if (_paragraphProperties != null && _paragraphProperties.NumberingProperties != null) {
                    return _paragraphProperties.NumberingProperties.NumberingLevelReference.Val;
                } else {
                    return null;
                }
            }
            set {
                if (_paragraphProperties != null && _paragraphProperties.NumberingProperties != null) {
                    if (_paragraphProperties.NumberingProperties.NumberingLevelReference != null) {
                        _paragraphProperties.NumberingProperties.NumberingLevelReference.Val = value;
                    }
                } else {
                    // should throw?
                }
            }
        }

        internal int? _listNumberId {
            get {
                if (_paragraphProperties != null && _paragraphProperties.NumberingProperties != null) {
                    return _paragraphProperties.NumberingProperties.NumberingId.Val;
                } else {
                    return null;
                }
            }
        }


        public WordParagraphStyles? Style {
            get {
                if (_paragraphProperties != null && _paragraphProperties.ParagraphStyleId != null) {
                    return WordParagraphStyle.GetStyle(_paragraphProperties.ParagraphStyleId.Val);
                }

                return null;
            }
            set {
                if (value != null) {
                    if (_paragraphProperties == null) {
                        _paragraph.ParagraphProperties = new ParagraphProperties();
                    }
                    if (_paragraphProperties.ParagraphStyleId == null) {
                        _paragraphProperties.ParagraphStyleId = new ParagraphStyleId();
                    }
                    _paragraphProperties.ParagraphStyleId.Val = value.Value.ToStringStyle();
                }
            }
        }


        internal WordList _list;
        internal List<Run> _runs;
        internal Hyperlink _hyperlink;
        internal SimpleField _simpleField;
        internal BookmarkStart _bookmarkStart;
        internal readonly OfficeMath _officeMath;
        internal readonly SdtRun _stdRun;
        internal readonly DocumentFormat.OpenXml.Math.Paragraph _mathParagraph;

        /// <summary>
        /// Get or set a text within Paragraph
        /// </summary>
        public string Text {
            get {
                if (_text == null) {
                    return "";
                }

                return _text.Text;
            }
            set {
                VerifyText();
                _text.Text = value;
            }
        }

        /// <summary>
        /// Get PageBreaks within Paragraph
        /// </summary>
        public WordBreak PageBreak {
            get {
                if (_run != null) {
                    var brake = _run.ChildElements.OfType<Break>().FirstOrDefault();
                    if (brake != null && brake.Type != null && brake.Type.Value == BreakValues.Page) {
                        return new WordBreak(_document, _paragraph, _run);
                    }
                }

                return null;
            }
        }

        /// <summary>
        /// Get Breaks within Paragraph
        /// </summary>
        public WordBreak Break {
            get {
                if (_run != null) {
                    var brake = _run.ChildElements.OfType<Break>().FirstOrDefault();
                    if (brake != null) {
                        return new WordBreak(_document, _paragraph, _run);
                    }
                }

                return null;
            }
        }

        public WordParagraph(WordDocument document = null, bool newParagraph = true, bool newRun = true) {
            this._document = document;

            if (newParagraph) {

                this._paragraph = new Paragraph();
                this._paragraph.AppendChild(new ParagraphProperties());

                if (newRun) {
                    this._run = new Run();
                    this._paragraph.AppendChild(_run);
                }
            }
        }

        public WordParagraph(WordDocument document, Paragraph paragraph) {
            this._document = document;
            this._paragraph = paragraph;
        }

        public WordParagraph(WordDocument document, Paragraph paragraph, Run run) {
            _document = document;
            _paragraph = paragraph;
            _run = run;
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, Hyperlink hyperlink) {
            _document = document;
            _paragraph = paragraph;
            _hyperlink = hyperlink;

            //this.Hyperlink = new WordHyperLink(document, paragraph, hyperlink);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, List<Run> runs) {
            _document = document;
            _paragraph = paragraph;
            _runs = runs;
            //this.Field = new WordField(document, paragraph, runs);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, SimpleField simpleField) {
            _document = document;
            _paragraph = paragraph;

            _simpleField = simpleField;

            //  this.Field = new WordField(document, paragraph, simpleField);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, BookmarkStart bookmarkStart) {
            _document = document;
            _paragraph = paragraph;

            _bookmarkStart = bookmarkStart;

            // this.Bookmark = new WordBookmark(document, paragraph, bookmarkStart);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
            _document = document;
            _paragraph = paragraph;

            _officeMath = officeMath;

            //this.Equation = new WordEquation(document, paragraph, officeMath);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, SdtRun stdRun) {
            _document = document;
            _paragraph = paragraph;
            _stdRun = stdRun;
            //this.StructuredDocumentTag = new WordStructuredDocumentTag(document, paragraph, stdRun);
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            _document = document;
            _paragraph = paragraph;
            _mathParagraph = mathParagraph;
            //  this.Equation = new WordEquation(document, paragraph, mathParagraph);
        }

        internal WordStructuredDocumentTag StructuredDocumentTag {
            get {
                if (_stdRun != null) {
                    return new WordStructuredDocumentTag(_document, _paragraph, _stdRun);
                }

                return null;
            }
        }

        public WordBookmark Bookmark {
            get {
                if (_bookmarkStart != null) {
                    return new WordBookmark(_document, _paragraph, _bookmarkStart);
                }

                return null;
            }
        }

        public WordEquation Equation {
            get {
                if (_officeMath != null || _mathParagraph != null) {
                    return new WordEquation(_document, _paragraph, _officeMath, _mathParagraph);
                }

                return null;
            }
        }

        public WordField Field {
            get {
                if (_simpleField != null || _runs != null) {
                    return new WordField(_document, _paragraph, _simpleField, _runs);
                }

                return null;
            }
        }

        public WordHyperLink Hyperlink {
            get {
                if (_hyperlink != null) {
                    return new WordHyperLink(_document, _paragraph, _hyperlink);
                }

                return null;
            }
        }

        public bool IsHyperLink {
            get {
                if (this.Hyperlink != null) {
                    return true;
                }

                return false;
            }
        }

        public bool IsField {
            get {
                if (this.Field != null && this.Field.Field != null) {
                    return true;
                }

                return false;
            }
        }

        public bool IsBookmark {
            get {
                if (this.Bookmark != null && this.Bookmark.Name != null) {
                    return true;
                }

                return false;
            }
        }

        public bool IsEquation {
            get {
                if (this.Equation != null) {
                    return true;
                }

                return false;
            }
        }

        public bool IsStructuredDocumentTag {
            get {
                if (this.StructuredDocumentTag != null) {
                    return true;
                }

                return false;
            }
        }

        public bool IsImage {
            get {
                if (this.Image != null) {
                    return true;
                }

                return false;
            }
        }

        public List<WordTab> TabStops {
            get {
                List<WordTab> list = new List<WordTab>();
                if (_paragraph != null && _paragraphProperties != null) {
                    if (_paragraphProperties.Tabs != null) {
                        foreach (TabStop tab in _paragraphProperties.Tabs) {
                            list.Add(new WordTab(this, tab));
                        }
                    }
                }
                return list;
            }
        }
    }
}
