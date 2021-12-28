using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;
using DocumentFormat.OpenXml;

namespace OfficeIMO {
    public partial class WordParagraph {
        internal WordDocument _document = null;
        internal Paragraph _paragraph = null;
        internal RunProperties _runProperties = null;
        internal Text _text = null;
        internal Run _run = null;
        internal ParagraphProperties _paragraphProperties;

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

        public WordParagraph(string text) {
            WordParagraph paragraph = new WordParagraph(this._document);
            paragraph.Text = text;
        }


        //public List<WordParagraph> GetParagraphs(List<Paragraph> list)
        //{
        //    var listWord = new List<WordParagraph>();
        //    //var list = this._wordprocessingDocument.MainDocumentPart.Document.Body.ChildElements.OfType<Paragraph>().ToList();
        //    foreach (Paragraph paragraph in list)
        //    {

        //        WordParagraph wordParagraph = new WordParagraph();

        //        //listWord.Add(wordParagraph);
        //        // foreach (var element in paragraph.ChildElements.OfType<Run>())
        //        // {
        //        //     
        //        //    }
        //    }

        //    return listWord;
        //}

        public WordParagraph(WordDocument document, Paragraph paragraph) {
            //_paragraph = paragraph;
            int count = 0;
            foreach (var run in paragraph.ChildElements.OfType<Run>()) {
                RunProperties runProperties = run.RunProperties;
                Text text = run.ChildElements.OfType<Text>().First();

                if (count > 0) {
                    WordParagraph wordParagraph = new WordParagraph(this._document);
                    wordParagraph._run = run;
                    wordParagraph._text = text;
                    wordParagraph._paragraph = paragraph;
                    wordParagraph._runProperties = runProperties;
                    document.Paragraphs.Add(wordParagraph);
                } else {
                    this._run = run;
                    this._text = text;
                    this._paragraph = paragraph;
                    this._runProperties = runProperties;
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
            return paragraph;
        }
        public WordParagraph InsertImage(string filePathImage) {
            WordImage wordImage = new WordImage(this._document, filePathImage, null, null);
            WordParagraph paragraph = new WordParagraph(this._document);
            _run.Append(wordImage._Image);
            return paragraph;
        }
    }
}