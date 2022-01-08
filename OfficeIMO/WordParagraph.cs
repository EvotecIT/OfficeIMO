using System;
using System.Collections.Generic;
using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Office2013.Drawing.TimeSlicer;

namespace OfficeIMO {
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
                if (_paragraphProperties.NumberingProperties != null) {
                    return true;
                } else {
                    return false;
                }
            }
        }

        internal WordList _list;

        public string Text {
            get { return _text.Text; }
            set { _text.Text = value; }
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
                document._currentSection.Paragraphs.Add(this);
                //document.Paragraphs.Add(this);
            }
        }

        public WordParagraph(WordDocument document, bool newParagraph, ParagraphProperties paragraphProperties, RunProperties runProperties, Run run) {
            this._document = document;
            this._run = run;
            this._runProperties = runProperties;
            this._text = new Text {
                // this ensures spaces are preserved between runs
                Space = SpaceProcessingModeValues.Preserve
            };
            this._paragraphProperties = paragraphProperties;
            this._run.AppendChild(_runProperties);
            this._run.AppendChild(_text);
            if (newParagraph) {
                this._paragraph = new Paragraph();
                this._paragraph.AppendChild(_paragraphProperties);
                this._paragraph.AppendChild(_run);
            }

            if (document != null) {
                document._currentSection.Paragraphs.Add(this);
                //document.Paragraphs.Add(this);
            }
        }

        public WordParagraph(string text) {
            WordParagraph paragraph = new WordParagraph(this._document);
            paragraph.Text = text;
        }

        /// <summary>
        /// Builds paragraph list when loading from filesystem
        /// </summary>
        /// <param name="document"></param>
        /// <param name="paragraph"></param>
        public WordParagraph(WordDocument document, Paragraph paragraph) {
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
                        wordParagraph._run = run;
                        wordParagraph._text = text;
                        wordParagraph._paragraph = paragraph;
                        wordParagraph._runProperties = runProperties;

                        wordParagraph.Image = newImage;

                        //document.Paragraphs.Add(wordParagraph);
                        document._currentSection.Paragraphs.Add(wordParagraph);
                        if (wordParagraph.IsPageBreak) {
                            //document.PageBreaks.Add(wordParagraph);
                            document._currentSection.PageBreaks.Add(wordParagraph);
                        }
                    } else {
                        this._run = run;
                        this._text = text;
                        this._paragraph = paragraph;
                        this._runProperties = runProperties;

                        if (newImage != null) {
                            this.Image = newImage;
                        }
                        //document.Paragraphs.Add(this);
                        //document._currentSection.Paragraphs.Add(this);

                        document._currentSection.Paragraphs.Add(this);
                        if (this.IsPageBreak) {
                            //document.PageBreaks.Add(this);
                            document._currentSection.PageBreaks.Add(this);
                        }
                    }

                    count++;
                }
            } else {
                // this is an empty paragraph so we add it
                document._currentSection.Paragraphs.Add(this);
            }
        }

        public WordParagraph InsertText(string text) {
            //DocumentFormat.OpenXml.Wordprocessing.Run run = new DocumentFormat.OpenXml.Wordprocessing.Run();
            //DocumentFormat.OpenXml.Wordprocessing.Text textProp = new DocumentFormat.OpenXml.Wordprocessing.Text();
            //textProp.Text = text;
            //run.AppendChild(textProp);
            //this.paragraph.Append(run);
            //return this;

            // WordParagraph paragraph = new WordParagraph();
            //this.paragraph.Text = text;
            //return this.paragraph;
            this._text.Text = text;
            return this;
        }

        public WordParagraph AppendText(string text) {
            WordParagraph wordParagraph = new WordParagraph(this._document, false);
            wordParagraph.Text = text;
          
            // this ensures that we keep track of matching runs with real paragraphs
            wordParagraph._linkedParagraph = this;

            if (this._linkedParagraph != null) {
                this._linkedParagraph._paragraph.Append(wordParagraph._run);
            } else {
                this._paragraph.Append(wordParagraph._run);
            }
            //this._document._wordprocessingDocument.MainDocumentPart.Document.InsertAfter(wordParagraph._run, this._paragraph);
            return wordParagraph;
        }
        
        public WordParagraph InsertImage(string filePathImage, double? width, double? height) {
            WordImage wordImage = new WordImage(this._document, filePathImage, width, height);
            WordParagraph paragraph = new WordParagraph(this._document);
            _run.Append(wordImage._Image);
            this.Image = wordImage;
            return paragraph;
        }
        public WordParagraph InsertImage(string filePathImage) {
            WordImage wordImage = new WordImage(this._document, filePathImage, null, null);
            WordParagraph paragraph = new WordParagraph(this._document);
            _run.Append(wordImage._Image);
            this.Image = wordImage;
            return paragraph;
        }

        public void Remove() {
            if (_paragraph != null) {
                if (this._paragraph.Parent != null) {
                    this._paragraph.Remove();
                } else {
                    throw new InvalidOperationException("This shouldn't happen? Why? Oh why?");
                    //Console.WriteLine(this._run);
                }
            } else {
                // this happens if we continue adding to real paragraphs additional runs. In this case we don't need to,
                // delete paragraph, but only remove Runs 
                this._run.Remove();
            }
            if (IsPageBreak) {
                this._document.PageBreaks.Remove(this);
            }

            if (IsListItem) {
                if (this._list != null) {
                    this._list.ListItems.Remove(this);
                    this._list = null;
                }
            }
            this._document.Paragraphs.Remove(this);
        }

        public WordParagraph InsertParagraphAfterSelf() {
            WordParagraph paragraph = new WordParagraph(null, true);
            this._paragraph.InsertAfterSelf(paragraph._paragraph);
            this._document.Paragraphs.Add(paragraph);
            
            return paragraph;
        }

        public WordParagraph InsertParagraphBeforeSelf() {
            WordParagraph paragraph = new WordParagraph(null, true);
            this._paragraph.InsertBeforeSelf(paragraph._paragraph);
            //document.Paragraphs.Add(paragraph);
            return paragraph;
        }
    }
}