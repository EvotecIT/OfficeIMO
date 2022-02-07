using System;
using System.Collections.Generic;
using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Word {
    public partial class WordParagraph {
        internal WordDocument _document;
        internal Paragraph _paragraph;
        internal RunProperties _runProperties;
        internal Text _text;
        internal Run _run;
        internal ParagraphProperties _paragraphProperties;
        internal WordParagraph _linkedParagraph;
        internal WordSection _section;

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
                        _paragraphProperties = new ParagraphProperties();
                    }

                    if (_paragraphProperties.ParagraphStyleId == null) {
                        _paragraphProperties.ParagraphStyleId = new ParagraphStyleId();
                    }

                    _paragraphProperties.ParagraphStyleId.Val = value.Value.ToStringStyle();
                }
            }
        }


        private WordList _list;
        private readonly Hyperlink _hyperlink;

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

        public WordPageBreak PageBreak {
            get {
                if (_run != null) {
                    var brake = _run.ChildElements.OfType<Break>().FirstOrDefault();
                    if (brake != null && brake.Type.Value == BreakValues.Page) {
                        return new WordPageBreak(_document, _paragraph, _run);
                    }
                }

                return null;
            }
        }

        public WordParagraph(WordSection section, bool newParagraph = true) {
            this._document = section._document;
            // this._section = section;

            this._run = new Run();
            this._runProperties = new RunProperties();

            //this._run = new Run();
            //this._runProperties = new RunProperties();
            this._text = new Text {
                // this ensures spaces are preserved between runs
                Space = SpaceProcessingModeValues.Preserve
            };
            this._paragraphProperties = new ParagraphProperties();
            this._run.AppendChild(_runProperties);
            this._run.AppendChild(_text);
            if (newParagraph) {
                this._paragraph = new Paragraph();
                this._paragraph.AppendChild(_paragraphProperties);
                this._paragraph.AppendChild(_run);
            }

            //section.Paragraphs.Add(this);
        }

        public WordParagraph(WordDocument document = null, bool newParagraph = true) {
            this._document = document;
            this._run = new Run();
            this._runProperties = new RunProperties();
            this._text = new Text {
                // this ensures spaces are preserved between runs
                Space = SpaceProcessingModeValues.Preserve
            };
            this._paragraphProperties = new ParagraphProperties();
            this._run.AppendChild(_runProperties);
            this._run.AppendChild(_text);
            if (newParagraph) {
                this._paragraph = new Paragraph();
                this._paragraph.AppendChild(_paragraphProperties);
                this._paragraph.AppendChild(_run);
            }

            if (document != null) {
                //  document._currentSection.Paragraphs.Add(this);
                //this._section = document._currentSection;
                //document.Paragraphs.Add(this);
            }
        }


        public WordParagraph(WordDocument document, Paragraph paragraph) {
            this._document = document;
            this._paragraph = paragraph;
        }

        public WordParagraph(WordDocument document, bool newParagraph, Paragraph paragraph, ParagraphProperties paragraphProperties, RunProperties runProperties, Run run, WordSection section = null) {
            this._document = document;
            this._section = section;
            this._run = run;
            this._runProperties = runProperties;
            this._paragraph = paragraph;

            if (run != null) this._text = run.OfType<Text>().FirstOrDefault();
            this._paragraphProperties = paragraphProperties;
            if (this._run != null) {
                //  this._run.AppendChild(_runProperties);
                // this._run.AppendChild(_text);
            }

            if (newParagraph) {
                this._paragraph = new Paragraph();
                if (_paragraphProperties != null) {
                    this._paragraph.AppendChild(_paragraphProperties);
                }
                if (_run != null) this._paragraph.AppendChild(_run);
            }
        }

        public WordParagraph(WordDocument document, Paragraph paragraph, Run run) {
            _document = document;
            _paragraph = paragraph;
            _run = run;
            if (run != null) {
                this._text = run.OfType<Text>().FirstOrDefault();

                if (run.RunProperties != null) {
                    _runProperties = run.RunProperties;
                }

                //Drawing drawing = run.ChildElements.OfType<Drawing>().FirstOrDefault();
                //if (drawing != null) {
                //    this.Image = new WordImage(document, drawing);
                //}
            }
            if (paragraph.ParagraphProperties != null) {
                _paragraphProperties = paragraph.ParagraphProperties;
            }
        }

        internal WordParagraph(WordDocument document, Paragraph paragraph, Hyperlink hyperlink) {
            _document = document;
            _paragraph = paragraph;
            _hyperlink = hyperlink;

            this.Hyperlink = new WordHyperLink(document, paragraph, hyperlink);
        }

        public WordParagraph(WordDocument document, Paragraph paragraph, List<Run> runs) {
            _document = document;
            _paragraph = paragraph;

            this.Field = new WordField(document, paragraph, runs);
        }

        public WordParagraph(WordDocument document, Paragraph paragraph, SimpleField simpleField) {
            _document = document;
            _paragraph = paragraph;

            this.Field = new WordField(document, paragraph, simpleField);
        }

        public WordParagraph(WordDocument document, Paragraph paragraph, BookmarkStart bookmarkStart) {
            _document = document;
            _paragraph = paragraph;

            this.Bookmark = new WordBookmark(document, paragraph, bookmarkStart);
        }

        public WordParagraph(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.OfficeMath officeMath) {
            _document = document;
            _paragraph = paragraph;

            this.Equation = new WordEquation(document, paragraph, officeMath);
        }

        public WordParagraph(WordDocument document, Paragraph paragraph, SdtRun stdRun) {
            _document = document;
            _paragraph = paragraph;

            this.StructuredDocumentTag = new WordStructuredDocumentTag(document, paragraph, stdRun);
        }

        public WordParagraph(WordDocument document, Paragraph paragraph, DocumentFormat.OpenXml.Math.Paragraph mathParagraph) {
            _document = document;
            _paragraph = paragraph;

            this.Equation = new WordEquation(document, paragraph, mathParagraph);
        }

        public WordStructuredDocumentTag StructuredDocumentTag { get; set; }

        public WordBookmark Bookmark { get; set; }

        public WordEquation Equation { get; set; }

        public WordField Field { get; set; }
        public WordHyperLink Hyperlink { get; set; }

        public bool IsHyperLink {
            get {
                if (this.Hyperlink != null && this.Hyperlink.Url != null) {
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
    }
}