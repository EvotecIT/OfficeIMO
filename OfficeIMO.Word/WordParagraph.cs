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

        public WordImage Image { get; set; }

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

        public WordParagraph(WordDocument document, bool newParagraph, Paragraph paragraph, ParagraphProperties paragraphProperties, RunProperties runProperties, Run run, WordSection section = null) {
            this._document = document;
            this._section = section;
            this._run = run;
            this._runProperties = runProperties;
            this._paragraph = paragraph;
            //this._text = new Text {
            //    // this ensures spaces are preserved between runs
            //    Space = SpaceProcessingModeValues.Preserve
            //};

            if (run != null) this._text = run.OfType<Text>().FirstOrDefault();
            this._paragraphProperties = paragraphProperties;
            if (this._run != null) {
                //  this._run.AppendChild(_runProperties);
                // this._run.AppendChild(_text);
            }

            if (newParagraph) {
                this._paragraph = new Paragraph();

                if (_paragraphProperties != null) {
                    //this._paragraph.ParagraphProperties = _paragraphProperties;
                    this._paragraph.AppendChild(_paragraphProperties);
                }
                if (_run != null) this._paragraph.AppendChild(_run);
            }

            if (document != null) {
                // document._currentSection.Paragraphs.Add(this);
                //section.Paragraphs.Add(this);
                //document.Paragraphs.Add(this);
            }
        }

        /// <summary>
        /// Used during loading of documents / tables only
        /// </summary>
        /// <param name="document"></param>
        /// <param name="paragraph"></param>
        /// <param name="section"></param>
        public WordParagraph(WordDocument document, Paragraph paragraph, WordSection section = null) {
            //_paragraph = paragraph;
            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.SectionProperties != null) {
                // TODO this means it's a section and we don't want to add sections to paragraphs don't we?

                this._paragraph = paragraph;
                return;
            }

            int count = 0;
            var listRuns = paragraph.ChildElements.OfType<Run>();
            if (listRuns.Any()) {
                foreach (var run in paragraph.ChildElements.OfType<Run>()) {
                    RunProperties runProperties = run.RunProperties;
                    Text text = run.ChildElements.OfType<Text>().FirstOrDefault();
                    Drawing drawing = run.ChildElements.OfType<Drawing>().FirstOrDefault();

                    WordImage newImage = null;
                    if (drawing != null) {
                        newImage = new WordImage(document, drawing);
                    }

                    if (count > 0) {
                        WordParagraph wordParagraph = new WordParagraph(this._document);
                        wordParagraph._document = document;
                        wordParagraph._run = run;
                        wordParagraph._text = text;
                        wordParagraph._paragraph = paragraph;
                        wordParagraph._paragraphProperties = paragraph.ParagraphProperties;
                        wordParagraph._runProperties = runProperties;
                        wordParagraph._section = section;

                        wordParagraph.Image = newImage;

                        //document._currentSection.Paragraphs.Add(wordParagraph);
                        if (wordParagraph.IsPageBreak) {
                            document._currentSection.PageBreaks.Add(wordParagraph);
                        }

                        if (wordParagraph.IsListItem) {
                            LoadListToDocument(document, wordParagraph);
                        }
                    } else {
                        this._document = document;
                        this._run = run;
                        this._text = text;
                        this._paragraph = paragraph;
                        this._paragraphProperties = paragraph.ParagraphProperties;
                        this._runProperties = runProperties;
                        this._section = section;

                        if (newImage != null) {
                            this.Image = newImage;
                        }

                        // this is to prevent adding Tables Paragraphs to section Paragraphs
                        if (section != null) {
                            section.Paragraphs.Add(this);
                            if (this.IsPageBreak) {
                                section.PageBreaks.Add(this);
                            }
                        }

                        if (this.IsListItem) {
                            LoadListToDocument(document, this);
                        }
                    }

                    count++;
                }
            } else {
                // this is an empty paragraph so we add it
                document._currentSection.Paragraphs.Add(this);
                this._section = document._currentSection;
            }
        }


    }
}