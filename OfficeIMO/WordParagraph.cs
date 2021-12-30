using System.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2013.Drawing.TimeSlicer;

namespace OfficeIMO {
    public partial class WordParagraph {
        internal WordDocument _document = null;
        internal Paragraph _paragraph = null;
        internal RunProperties _runProperties = null;
        internal Text _text = null;
        internal Run _run = null;
        internal ParagraphProperties _paragraphProperties;

        public WordImage Image { get; set; }

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
                document.Paragraphs.Add(this);
            }
        }

        public WordParagraph() {

        }
        public WordParagraph(string text) {
            WordParagraph paragraph = new WordParagraph(this._document);
            paragraph.Text = text;
        }
        public WordParagraph(WordDocument document, Paragraph paragraph) {
            //_paragraph = paragraph;
            int count = 0;
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

                    document.Paragraphs.Add(wordParagraph);
                } else {
                    this._run = run;
                    this._text = text;
                    this._paragraph = paragraph;
                    this._runProperties = runProperties;

                    if (newImage != null) {
                        this.Image = newImage;
                    }

                    document.Paragraphs.Add(this);
                }

           

                count++;
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
            this._paragraph.Append(wordParagraph._run);
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
            this._paragraph.Remove();
            this._document.Paragraphs.Remove(this);
        }
    }
}